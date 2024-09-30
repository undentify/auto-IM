
# Python Script REST API ADCM

import requests
import json


ip_adcm = "10.128.0.90"

# Get Token
response = requests.post('http://' + ip_adcm + ':8000/api/v2/token/', data={'username': 'admin', 'password': 'admin'})
token = response.json()['token']
print('Token is: ',token)
# List hosts
headers = {"Authorization": f'Token {token}'}
response = requests.get('http://' + ip_adcm + ':8000/api/v2/hosts/', headers=headers)
#print(response.json())

# List of hosts
list_of_hosts = [
    "adhadm-999-adcm.ru-central1.internal"
#    "adhadm-111-mnode01.ru-central1.internal",
#    "adhadm-111-mnode02.ru-central1.internal",
#    "adhadm-111-snode03.ru-central1.internal",
#    "adhadm-111-snode04.ru-central1.internal",
#    "adhadm-111-snode05.ru-central1.internal",
#   "adhadm-111-snode06.ru-central1.internal"
]

# Create hosts with username/password/hostname configuration
print('creating hosts:',list_of_hosts)
url = 'http://' + ip_adcm + ':8000/api/v2/hosts/'
headers = {"Authorization": f'Token {token}'}
for host in list_of_hosts:
    response = requests.post(
        url,
        headers=headers,
        data={"ansible_user": "admin",
              # "prototype_id": "81",
              # prototype for host object api/v2/prototypes/
              "hostprovider_id": "1",  # id provider /api/v1/provider/
              "name": host})
    if response.status_code is not 201:
        print(response.json())

    if#response.status_code == 201:
        url_host_history = url + str(response.json()['id']) + "/configs/"
        response_history = requests.post(url_host_history, headers=headers, data={
            "config": json.dumps({
                                "__main_info": "python test",
                                "ansible_user": "admin",
                                "ansible_ssh_pass": "adh@2015",
                                "ansible_host": host,
                                "ansible_ssh_port": "22",
                                "ansible_ssh_common_args": "-o StrictHostKeyChecking=no -o UserKnownHostsFile=/dev/null",
                                "ansible_become": bool('true')}),
            "adcmMeta": json.dumps({}),
            'description': "python test desc"
        }
                                         )
#        print(response_history.status_code)
        print('Host ',host,' created, config_id is: ',response_history.json()['id'])




# # Upload bundle
# url = 'http://' + ip_adcm + ':8000/api/v1/stack/upload/'
# files = {'file': open('c:/temp/adcm_cluster_hadoop_v2.1.4_b6-1_enterprise.tgz', 'rb')}
# headers = {"Authorization": f'Token {token}'}
# response = requests.post(url, headers=headers, files=files)
# print(response.status_code)
# print(response.content)

# Load bundle
# url = 'http://' + ip_adcm + ':8000/api/v1/stack/load/'
# headers = {"Authorization": f'Token {token}'}
# response = requests.post(url, headers=headers, data={'bundle_file': 'adcm_cluster_hadoop_v2.1.4_b6-1_enterprise.tgz'})
# print(response.status_code)
# print(response.content)


