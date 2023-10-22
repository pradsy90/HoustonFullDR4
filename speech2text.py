import azure.cognitiveservices.speech as speechsdk
import os
from azure.cognitiveservices.speech import AudioConfig,ResultFuture
import re


def from_mic():
    audio_config = AudioConfig(device_name="<device id>");

    speech_config = speechsdk.SpeechConfig(subscription=os.environ.get('SPEECH_KEY'), region=os.environ.get('SPEECH_REGION'))
    speech_recognizer = speechsdk.SpeechRecognizer(speech_config=speech_config)
    #print("Speak into your microphone.")
    result =speech_recognizer.recognize_once_async().get()
    print(result.text)
    return(process_results(str(result.text)))


def process_results(x):
    x=x.replace("Task","")
    print(x)
    x=x.replace("comma",",")
    print(x)
    x=x.replace("Tap","")
    print(x)
    x=x.replace("Tab","")
    print(x)
    x=x.replace("task","")
    print(x)
    x=x.replace("tap","")
    print(x)
    x=x.replace("app","")
    print(x)
    x=x.replace("and","")
    print(x)
    x=x.replace(".","")
    print(x)
    x=x.replace(" ","")
    print(x)
    x=x.replace(",","")
    print(x)
    x = re.findall("[0-9]{5}", x)
    for tasknumber in x:
        print (tasknumber + " IS THE ITEM")
        print (len(tasknumber))
    return x