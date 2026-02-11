
import json
import requests
import urllib.parse

def update_case(case_id, token):

    response = requests.get(f"https://vejman.vd.dk/permissions/getcase?caseid={case_id}&token={token}")
    response.raise_for_status()
    json_object = response.json().get('data')

    filtered_data = {k: json_object[k] for k in [
    "type", "variant", "origin", "state", "year", "serial_number", "authority_reference_number", 
    "start_date", "end_date", "initials", "visuser_id", "created_date", "created_user", 
    "modified_date", "modified_user", "connected_case", "bestyrer", "community", 
    "majorVersion", "minorVersion", "authName", "authEmail", "case_set", "brokerCaseState", "id"
    ] if k in json_object}
    # Add "$transaction" and "$changed" nodes
    filtered_data["authority_reference_number"] = "Faktura sendt"
    filtered_data["$transaction"] = "update"
    filtered_data["$changed"] = True

    # Convert the dictionary to a compact JSON string without spaces, ensuring UTF-8 encoding
    json_data = json.dumps(filtered_data, ensure_ascii=False, separators=(',', ':'))

    # URL-encode the JSON string
    url_encoded_data = urllib.parse.quote(json_data)

    # Construct the payload string for the POST request
    payload = f"data={url_encoded_data}"

    
    # Make the POST request with the encoded data
    post_url = f"https://vejman.vd.dk/permissions/setcase?token={token}"
    headers = {
        'Accept': 'text/javascript, text/html, application/xml, text/xml, */*',
        'Content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    }

    response = requests.post(post_url, headers=headers, data=payload)
    response.raise_for_status()

    post_response_data = response.json()

    # Check if 'data' key exists in response and compare 'id'
    if 'data' in post_response_data and post_response_data['data'].get('id') == filtered_data.get('id'):
        print(f"Case ID {filtered_data['id']}: Data updated successfully!")
    else:
        print(f"Case ID {filtered_data['id']}: Failed to update data, no matching ID found.")