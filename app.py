import streamlit as st
import msal
import requests
import json
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

def fetch_emails(access_token, user_email):
    """Fetch emails from Outlook with metadata."""
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages?$select=subject,from,body,receivedDateTime"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json().get("value", [])
    else:
        st.error(f"Error fetching emails: {response.status_code} - {response.text}")
        return []

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
    topics = extract_topics(mails)
    relevant_mails = [mail for mail in mails if any(topic in mail.get('subject', '') for topic in topics)]
    
    if not relevant_mails:
        return "No relevant emails found."
    
    mail_details = "\n".join([
        f"Subject: {mail.get('subject', 'No Subject')}\nBody: {mail.get('body', {}).get('content', '')}"
        for mail in relevant_mails[:10]
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
st.title("Outlook Mail Viewer with Smart Filtering")  
user_email = st.text_input("Enter User Email")  
user_query = st.text_input("Ask a question about the emails")

if st.button("Ask"):
    token = get_access_token()
    if token and user_email:
        mails = fetch_emails(token, user_email)
        st.write(f"Found {len(mails)} email(s)")
        
        if user_query:
            answer = query_responder(user_query, mails)
            st.write(f"Answer: {answer}")
    else:
        st.error("Invalid email or authentication issue.")
