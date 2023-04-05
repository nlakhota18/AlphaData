import openai
import pandas as pd
import requests
openai.api_key = "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"


audio_file = open("C:\\AlphaData\\ChatGPT Calls\\20230404-141322_4702525044-all.mp3", "rb")
transcript = openai.Audio.transcribe("whisper-1", audio_file, speaker_labels=True)




# Analyze text using OpenAI API
def analyze_text(transcript):
    model_engine = "text-davinci-002"
    
    # Modify prompt to ask a direct yes/no question and provide more context
    prompt = (
        f"Given the following transcript of a call between an agent and a customer. The purpose of the call is to ask the customer if they would like to receive testing kits for free from medicare.:\n{transcript}\n"
        "Did the customer consent to receiving test kits from Alpha Labs or consent to receiving test kits anywhere in the call?","Did the customer explicitly say 'No' to receiving kits?"
    )
    
    completions = openai.Completion.create(
        engine=model_engine,
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.5,
    )
    message = completions.choices[0].text.strip()
    return message


print(transcript)
message = analyze_text(transcript)
print(message)
