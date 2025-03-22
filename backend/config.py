import os
from dotenv import load_dotenv

# Load biến môi trường từ file .env
load_dotenv()

class Config:
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "your_default_api_key_here")

    HOST = "0.0.0.0"
    PORT = 5000
    DEBUG = True  
    OPENAI_MODEL = "gpt-4"
    DEFAULT_TEMPERATURE = 0.7  

    MAX_TEXT_LENGTH = 5000  

config = Config()
