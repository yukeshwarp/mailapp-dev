import streamlit as st
import msal
import requests
import os
from openai import AzureOpenAI
import html2text
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import NMF

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
    """Extract relevant topics using NMF."""
    h = html2text.HTML2Text()
    h.ignore_links = True
    
    mail_texts = [
        f"Subject: {mail.get('subject', 'No Subject')}\nBody: {h.handle(mail.get('body', {}).get('content', ''))}"
        for mail in mails
    ]
    
    vectorizer = TfidfVectorizer(stop_words="english", max_features=1000)
    tfidf = vectorizer.fit_transform(mail_texts)
    
    n_topics = min(max_topics, tfidf.shape[1])
    nmf = NMF(n_components=n_topics, random_state=42, max_iter=500)
    nmf.fit(tfidf)
    
    feature_names = vectorizer.get_feature_names_out()
    topics = [
        ", ".join([feature_names[i] for i in topic.argsort()[-max_top_words:][::-1]])
        for topic in nmf.components_
    ]
    return topics

def query_responder(query, mails):
    """Use LLM to respond to user query based on filtered emails."""
    h = html2text.HTML2Text()  
    h.ignore_links = True  
    
    topics = extract_topics(mails)
    mail_details = "\n".join([
        f"Subject: {mail.get('subject', 'No Subject')}\n"
        f"From: {mail.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender')}\n"
        f"Received: {mail.get('receivedDateTime', 'Unknown Time')}\n"
        f"Body: {h.handle(mail['body']['content']) if mail.get('body', {}).get('contentType') == 'html' else mail.get('body', {}).get('content', 'No Content')}"
        for mail in mails 
    ])
    
    prompt = f"Answer the user's query using these emails:\n{mail_details}\n\nUser's Query: {query}"
    
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.5,
    )
    return response.choices[0].message.content.strip()

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
        response = query_responder(prompt, mails)
    else:
        response = "No emails available. Fetch emails first."
    
    st.session_state["messages"].append({"role": "assistant", "content": response})
    with st.chat_message("assistant"):
        st.markdown(response)
