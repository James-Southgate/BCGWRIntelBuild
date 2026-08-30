#!/usr/bin/python

__author__ = "HPInc."

import requests, hmac, hashlib, datetime
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
from urllib.parse import unquote, urlsplit


class OneflowSDK:
    def __init__(self, url, token, secret, options=None):
        self.url = url
        self.token = token
        self.secret = secret
        self.options = options

    def requests_retry_session(self):
        retries = getattr(self.options, "retries", 3)
        retryFactor = getattr(self.options, "retryFactor", 0.3 * 60)
        retryStatus = getattr(self.options, "retryStatus", [429])
        retryMethods = getattr(
            self.options,
            "retryMethods",
            ["HEAD", "GET", "PUT", "DELETE", "OPTIONS", "TRACE", "POST"],
        )

        session = requests.Session()
        retry = Retry(
            total=retries,
            read=retries,
            connect=retries,
            backoff_factor=retryFactor,
            status_forcelist=retryStatus,
            allowed_methods=retryMethods,
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        return session

    def create_headers(self, method, path):
        timestamp = (
            datetime.datetime.now(datetime.timezone.utc)
            .isoformat(timespec="milliseconds")
            .replace("+00:00", "Z")
        )

        path_without_query = urlsplit(path).path

        decoded_path = unquote(path_without_query)

        string_to_sign = (
            f"{method.upper()} {decoded_path} {timestamp}"
        )

        signature = hmac.new(
            self.secret.encode("utf-8"),
            string_to_sign.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()

        return {
            "content-type": "application/json",
            "x-oneflow-algorithm": "SHA256",
            "x-oneflow-authorization": f"{self.token}:{signature}",
            "x-oneflow-date": timestamp,
        }

    def request(self, method, path, data=None, params=None):
        url = self.url + path
        headers = self.create_headers(method, path)
        session = self.requests_retry_session()
        result = session.request(method, url, data=data, params=params, headers=headers)
        return result.content