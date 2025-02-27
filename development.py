import streamlit as st
import msal
import requests
import os
from openai import AzureOpenAI
import html2text
import json
import re
from datetime import datetime, timedelta
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import Normalizer
from sklearn.decomposition import NMF


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
    """Uses LLM to extract start and end dates from the query."""  
    prompt = f"""  
    You are a smart assistant tasked with identifying date ranges from user queries. Here's what you need to do:

    Analyze the user query to find any mentioned dates or date ranges.
    If a specific date is mentioned, set both start_date and end_date to that date.
    If relative dates are used (e.g., "last month", "next week"), calculate the absolute dates based on today's date.
    If no dates are mentioned, return null for both start_date and end_date.

    Provide the output in the following JSON format:

        "start_date": "YYYY-MM-DD",  
        "end_date": "YYYY-MM-DD"  
    
    User Query: {query}
    """  
  
    response = client.chat.completions.create(  
        model="gpt-4o",  
        messages=[  
            {"role": "system", "content": "You extract date ranges from queries."},  
            {"role": "user", "content": prompt},  
        ],  
        temperature=0,  
    )  
  
    # Extract JSON from LLM response  
    json_match = re.search(r"\{.*?\}", response.choices[0].message.content.strip(), re.DOTALL)  
    if json_match:  
        try:  
            date_info = json.loads(json_match.group(0))  
            return date_info.get('start_date'), date_info.get('end_date')  
        except json.JSONDecodeError:  
            st.error("Error parsing dates from LLM response.")  
            return None, None  
  
    st.error("No date information found in LLM response.")  
    return None, None  

def filter_mails_by_date(mails, start_date, end_date):  
    """Filters emails that fall within the given start_date and end_date."""  
    if not start_date or not end_date:  
        return mails  # If no valid date found, return all emails  
  
    start_datetime = datetime.strptime(start_date, "%Y-%m-%d")  
    end_datetime = datetime.strptime(end_date, "%Y-%m-%d")  
  
    return [  
        mail for mail in mails  
        if "receivedDateTime" in mail and  
        start_datetime <= datetime.strptime(mail["receivedDateTime"][:19], "%Y-%m-%dT%H:%M:%S") <= end_datetime  
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
