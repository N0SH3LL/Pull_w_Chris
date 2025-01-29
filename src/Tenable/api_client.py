"""Honestly pretty simple to interact with Tenable's API. more info here: https://developer.tenable.com/reference/navigate """

import os
import requests # requests lib does the heavy lifting here 
import json

class TenableSCClient: # base class for API requets
    def __init__(self, access_key:str, secret_key:str):
        self.base_url = "https://tenable-sc.corp.wpsic.com/rest/"
        self.access_key = access_key # GUI prompts for these, so they never touch disk
        self.secret_key = secret_key
        self.headers = {
            'X-ApiKey': f'accessKey={self.access_key}; secretKey={self.secret_key}',
            'Content-Type': 'application/json'
        }

        self.ca_cert_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'config', 'cert.cer'))
    # GET
    def get(self, endpoint, params=None):
        url = f"{self.base_url}/{endpoint}"
        response = requests.get(url, headers=self.headers, params=params, verify=self.ca_cert_path)
        response.raise_for_status()
        return response.json()
    # POST
    def post(self, endpoint, data, raw_response=False):
        url = f"{self.base_url}/{endpoint}"
        response = requests.post(url, headers=self.headers, json=data, verify=self.ca_cert_path)
        response.raise_for_status()
        if raw_response:
            return response
        return response.json()
    # PATCH 
    def patch(self, endpoint, data):
        url = f"{self.base_url}/{endpoint}"
        response = requests.patch(url, headers=self.headers, json=data, verify=self.ca_cert_path)
        response.raise_for_status()
        return response.json()
   
