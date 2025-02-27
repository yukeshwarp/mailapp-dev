import streamlit as st
import msal
import requests
import os
from openai import AzureOpenAI
import html2text
import time
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import Normalizer
from sklearn.decomposition import NMF
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import json  # Import JSON module for safe parsing
import re

# Azure app registration details
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# LLM setup
client = AzureOpenAI(
    azure_endpoint=os.getenv("LLM_ENDPOINT"),
    api_key=os.getenv("LLM_KEY"),
    api_version="2024-10-01-preview",
)

def get_access_token():  
    """Authenticate and get access token."""
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in result:
        return result["access_token"]
    else:
        st.error(f"Error acquiring token: {result.get('error_description')}")
        return None

@st.cache_data(show_spinner=True)
def fetch_emails(access_token, user_email):
    """Fetch all emails from Outlook with metadata."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    all_mails = []
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            all_mails.extend(data.get("value", []))
            url = data.get("@odata.nextLink")  # Get next page URL if available
        else:
            st.error(f"Error fetching emails: {response.status_code} - {response.text}")
            break

    return all_mails


def extract_topics(mails, max_topics=5, max_top_words=10):
    """Extract relevant topics using NMF and TF-IDF."""
    h = html2text.HTML2Text()
    h.ignore_links = True
    
    # Extract email text (subject + body)
    mail_texts = [
        f"Subject: {mail.get('subject', 'No Subject')}\nBody: {h.handle(mail.get('body', {}).get('content', ''))}"
        for mail in mails if mail.get("body", {}).get("content", "").strip()  # Ignore empty emails
    ]
    
    if not mail_texts:
        return []  # No meaningful data

    # TF-IDF Vectorization
    vectorizer = TfidfVectorizer(stop_words="english", max_features=1000)
    tfidf = vectorizer.fit_transform(mail_texts)

    # Normalize the TF-IDF matrix for better NMF performance
    normalizer = Normalizer(copy=False)
    tfidf = normalizer.fit_transform(tfidf)

    # NMF Model
    n_topics = min(max_topics, tfidf.shape[1])  # Prevent overfitting if features are fewer
    nmf = NMF(n_components=n_topics, random_state=42, max_iter=500, init="nndsvd")
    nmf.fit(tfidf)

    # Extract topic words
    feature_names = vectorizer.get_feature_names_out()
    topics = [
        ", ".join([feature_names[i] for i in topic.argsort()[-max_top_words:][::-1]])
        for topic in nmf.components_
    ]
    return topics


def fetch_relevant_mails(mails, query):
    """Use LLM to fetch relevant email IDs based on the query."""
    if not mails:
        return []

    # Extract relevant metadata
    mail_details = [
        {
            "Email ID": mail.get("id", "Unknown"),
            "From": mail.get("from", {}).get("emailAddress", {}).get("address", "Unknown"),
            "Received": mail.get("receivedDateTime", "Unknown"),
            "Body Preview": mail.get("bodyPreview", "No Preview"),
        }
        for mail in mails
    ]

    # Generate LLM prompt
    prompt = f"""
    Identify the most relevant email IDs based on the user's query.

    User Query: {query}

    Emails:
    {mail_details}

    Return ONLY a JSON array of email IDs. Example:
    ["AAMkADk5YzNmMGUxLTk1ZWItNDdkNi04OGMyLWIyYjg4NWVhY2ZmYw", "AAMkBzIzNmMGUxLTk1ZWItNDdkNi04OGMyLWIyYjg4NWVhY2ZmYx"]
    """

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are an intelligent email assistant. Respond with JSON only."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.3,
    )

    relevant_email_ids = response.choices[0].message.content.strip()

    # Extract JSON using regex
    json_match = re.search(r"\[.*\]", relevant_email_ids, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(0))  # Safely parse extracted JSON
        except json.JSONDecodeError:
            st.error("Error: Extracted text is not valid JSON.")
            return []

    st.error("Error: No JSON array found in LLM response.")
    return []



# Streamlit UI
st.set_page_config(page_title="Outlook Mail Assistant", layout="wide")
st.sidebar.title("Email Input")
user_email = st.sidebar.text_input("Enter User Email")

if st.sidebar.button("Fetch Emails"):
    token = get_access_token()  # Replace with your authentication function
    if token and user_email:
        mails = fetch_emails(token, user_email)
        st.session_state["mails"] = mails
        st.sidebar.success(f"Fetched {len(mails)} emails")
    else:
        st.sidebar.error("Invalid email or authentication issue.")

# Chat Interface
st.title("Outlook Mail Chat Assistant")
if "messages" not in st.session_state:
    st.session_state["messages"] = []

for message in st.session_state["messages"]:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if prompt := st.chat_input("Ask a question about your emails"):
    st.session_state["messages"].append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    mails = st.session_state.get("mails", [])
    if mails:
        # response = query_responder(prompt, mails)
        if not mails:
            st.error("No emails available. Please fetch emails first.")
        """Use LLM to respond to a user query based on the most relevant emails."""
        if not mails:
            st.write("No emails available. Please fetch emails first.")
        
        relevant_email_ids = fetch_relevant_mails(mails, prompt)

        # Filter emails using 'id' instead of 'conversationId'
        relevant_mails = [mail for mail in mails if mail.get("id") in relevant_email_ids]

        
        if not relevant_mails:
            st.write("No relevant emails found.")
        
        h = html2text.HTML2Text()
        h.ignore_links = True
        
        mail_details = "\n".join([
            f"Subject: {mail.get('subject', 'No Subject')}\n"
            f"From: {mail.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender')}\n"
            f"Received: {mail.get('receivedDateTime', 'Unknown Time')}\n"
            f"Importance: {mail.get('importance', 'Normal')}\n"
            f"Has Attachment: {mail.get('hasAttachments', False)}\n"
            f"Categories: {', '.join(mail.get('categories', [])) if mail.get('categories') else 'None'}\n"
            f"Conversation ID: {mail.get('conversationId', 'N/A')}\n"
            f"Weblink: {mail.get('webLink', 'No Link')}\n"
            f"Body: {h.handle(mail['body']['content']) if mail.get('body', {}).get('contentType') == 'html' else mail.get('body', {}).get('content', 'No Content')}"
            for mail in relevant_mails[:25]
        ])
    
        # Generate LLM prompt
        with st.spinner("Thinking..."):
            prompt_template = f"Answer the user's query using these emails:\n\n" + mail_details + f"\n\nUser's Query: {prompt}"
        
            # Call LLM
            response_stream = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt_template},
                ],
                temperature=0.5,
                stream = True,
            )
        bot_response = ""
        with st.chat_message("assistant"):
            response_placeholder = st.empty()
            for chunk in response_stream:
                if chunk.choices:
                    bot_response += chunk.choices[0].delta.content or ""
                    response_placeholder.markdown(bot_response)


    else:
        response = "No emails available. Fetch emails first."
    
    st.session_state["messages"].append({"role": "assistant", "content": bot_response})
