import requests
from requests_toolbelt import MultipartEncoder
import os

folder_path = os.path.join(os.path.expanduser("~/Desktop"), "DARA BOT", "config", "DARA Generator", "output")
#file_name = "C:/Users/Acronis/Desktop/REPORT BOT - Copy/process/python/Book1.xls"
#user_id = "c46c266b"
#chat_id = "oc_03a48538f2135c6470dd564cb2dbb6fc" #TFS
#chat_id = "oc_875b5c127cc9bf4c8602a050a41c27f5" #RCBC
chat_id = "oc_1cbed65e2b7015df3939a5c1f55f96e5" #RCBC DARA REQUEST
#chat_id = "oc_f8b5da69bf1dc3a1844946205a4efabb" #Task Force Agila

def upload(folder_path, chat_id):

    ## Request Token
    request_token_url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    request_token_data = {
            "app_id": "cli_a2bffebdbd78d009",
            "app_secret": "hQBykhX6T8iLGmzzJGHzxh0rBE0iBAOF"
            }
    gettoken = requests.post(url=request_token_url, data=request_token_data)
    token = gettoken.json()['tenant_access_token']

    ## Upload all PDF file in folder
    url = "https://open.larksuite.com/open-apis/im/v1/files"
    for filename in os.listdir(folder_path):
        if os.path.splitext(filename)[1].lower() == ".pdf":
            file_path = os.path.join(folder_path, filename)
            form = {'file_type': 'pdf',
                        'file_name': os.path.basename(file_path),
                        'file':  (os.path.basename(file_path), open(file_path, 'rb'), 'application/pdf')} 
            multi_form = MultipartEncoder(form)
            headers = {
            'Authorization': 'Bearer '+token,
            }

            headers['Content-Type'] = multi_form.content_type
            response = requests.request("POST", url, headers=headers, data=multi_form)

            print(response.headers['X-Tt-Logid']) # for debug or oncall
            print(response.content) # Print Response

            ## Get File_key
            file_key = response.json()
            file_key = file_key['data']['file_key']
            print(file_key)

            ## Send Message
            send_message_url = "https://open.larksuite.com/open-apis/im/v1/messages"
            send_message_params = {"receive_id_type":"chat_id"}
            send_message_headers = {'Authorization' : f'Bearer {token}'}
            send_message_text = {
                "receive_id": chat_id,
                "title": "Title",
                "content": "{\"text\":\"Hello World\"}",
                "msg_type": "text"
            }

            send_message_file = {
                "receive_id": chat_id,
                "content": "{\"file_key\":\""+file_key+"\"}",
                "msg_type": "file"
            }

            #result = requests.post(url=send_message_url, params=send_message_params, headers=send_message_headers, data=send_message_text)
            result = requests.post(url=send_message_url, params=send_message_params, headers=send_message_headers, data=send_message_file)
            
            print(result.text)

upload(folder_path, chat_id)
