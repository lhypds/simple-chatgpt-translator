import itertools
import time
import openai
from abc import ABC, abstractmethod


class ChatGPTAPI(ABC):
    def __init__(self, key, language):
        self.keys = itertools.cycle(key.split(","))
        self.language = language
        self.key_len = len(key.split(","))


    def rotate_key(self):
        openai.api_key = next(self.keys)


    def translate(self, text):
        print("Original: " + text)
        self.rotate_key()
        try:
            completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "user",
                        # english prompt here to save tokens
                        "content": f"Please help me to translate,`{text}` to {self.language}, please return only translated content not include the origin text",
                    }
                ],
            )
            t_text = (
                completion["choices"][0]  # type: ignore
                .get("message")
                .get("content")
                .encode("utf8")
                .decode()
            )
        except Exception as e:
            # TIME LIMIT for open api please pay
            sleep_time = int(60 / self.key_len)
            time.sleep(sleep_time)
            print(e, f"will sleep {sleep_time} seconds")
            self.rotate_key()
            completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {
                        "role": "user",
                        "content": f"Please help me to translate,`{text}` to {self.language}, please return only translated content not include the origin text",
                    }
                ],
            )
            t_text = (
                completion["choices"][0]  # type: ignore
                .get("message")
                .get("content")
                .encode("utf8")
                .decode()
            )
        
        # trim the text
        t_text = t_text.strip()
        if "→" in t_text:
            t_text = t_text.split("→")[1]
        t_text = t_text.strip()
        
        print("Trasnation: " + t_text)
        return t_text
