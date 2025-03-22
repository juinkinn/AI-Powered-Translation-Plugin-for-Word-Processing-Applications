import requests

url = "http://127.0.0.1:5000/translate"
data = {
    "text": "Hello",
    "source_lang": "English",
    "target_lang": "Vietnamese"
}

response = requests.post(url, json=data)
print(response.json())  # In kết quả dịch
