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
    return result.get("access_token", None)

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

def extract_dates_from_query(query):
    """Extracts date range from the query and returns start & end date as YYYY-MM-DD."""
    today = datetime.utcnow()
    match = re.search(r"(\d{4})\s*(January|February|March|April|May|June|July|August|September|October|November|December)", query, re.IGNORECASE)
    
    if match:
        year = int(match.group(1))
        month = match.group(2).capitalize()
        month_number = datetime.strptime(month, "%B").month
        start_date = datetime(year, month_number, 1).strftime("%Y-%m-%d")
        end_date = (datetime(year, month_number + 1, 1) - timedelta(days=1)).strftime("%Y-%m-%d") if month_number < 12 else (datetime(year + 1, 1, 1) - timedelta(days=1)).strftime("%Y-%m-%d")
        return start_date, end_date
    
    return None, None

def filter_mails_by_date(mails, start_date, end_date):
    """Filters emails that fall within the given start_date and end_date."""
    if not start_date or not end_date:
        return mails  # If no valid date found, return all emails
    
    return [
        mail for mail in mails
        if "receivedDateTime" in mail and start_date <= mail["receivedDateTime"][:10] <= end_date
    ]

def fetch_relevant_mails(mails, query):
    """Filter emails based on query and LLM."""
    start_date, end_date = extract_dates_from_query(query)
    filtered_mails = filter_mails_by_date(mails, start_date, end_date)

    if not filtered_mails:
        return []

    h = html2text.HTML2Text()
    h.ignore_links = True

    mail_details = [
        {
            "Email ID": mail.get("id", "Unknown"),
            "From": mail.get("from", {}).get("emailAddress", {}).get("address", "Unknown"),
            "Received": datetime.strptime(mail["receivedDateTime"], "%Y-%m-%dT%H:%M:%SZ").strftime("%d %B %Y, %H:%M UTC"),
            "Subject": mail.get("subject", "No Subject"),
            "Body Preview": mail.get("bodyPreview", "No Preview"),
        }
        for mail in filtered_mails
    ]

    prompt = f"""
    Identify relevant emails and return their ID. If no specific emails match, return all emails.

    Query: {query}

    Emails:
    {json.dumps(mail_details, indent=2)}

    Return ONLY a JSON array of email IDs, like:
    ["AAMk123...", "AAMk456..."]
    """

    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": "You are an email sorting assistant. Respond in JSON only."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.3,
    )

    json_match = re.search(r"\[.*\]", response.choices[0].message.content.strip(), re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(0))
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
    token = get_access_token()
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
        relevant_email_ids = fetch_relevant_mails(mails, prompt)
        relevant_mails = [mail for mail in mails if mail.get("id") in relevant_email_ids]

        if not relevant_mails:
            st.write("No relevant emails found.")

        h = html2text.HTML2Text()
        h.ignore_links = True

        mail_details = "\n".join([
            f"Subject: {mail.get('subject', 'No Subject')}\n"
            f"From: {mail.get('from', {}).get('emailAddress', {}).get('address', 'Unknown Sender')}\n"
            f"Received: {mail.get('receivedDateTime', 'Unknown Time')}\n"
            f"Body: {h.handle(mail.get('body', {}).get('content', 'No Content'))}"
            for mail in relevant_mails[:25]
        ])

        with st.spinner("Thinking..."):
            response_stream = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "system", "content": "Answer the user's query based on the given emails."}, {"role": "user", "content": mail_details + f"\n\nUser's Query: {prompt}"}],
                temperature=0.5,
                stream=True,
            )
        
        bot_response = ""
        with st.chat_message("assistant"):
            response_placeholder = st.empty()
            for chunk in response_stream:
                if chunk.choices:
                    bot_response += chunk.choices[0].delta.content or ""
                    response_placeholder.markdown(bot_response)

    st.session_state["messages"].append({"role": "assistant", "content": bot_response})
