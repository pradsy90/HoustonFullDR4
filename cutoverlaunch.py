#stuff just for py2.exe
import site
from sys import executable
from os import path
#interpreter = executable
#sitepkg = path.dirname(interpreter) + "\\site-packages"
#site.addsitedir(sitepkg)

#from threading import current_thread
#from tkinter import messagebox
import re
import PySimpleGUI as sg
import azure.cognitiveservices.speech as speechsdk
import recognizevoice
import scantask
import texttospeech
import base64
from datetime import datetime
import os
import multiprocessing
#from multiprocessing import Process, current_process
import pyautogui
import time
import speech2text

#Basic Variables to use in all functions
recbuttonpressed=False
sg.theme('LightGrey5')
#places where icons will be located
recordingfilename="C:\\Cutover\\voicerecord1_png.png"
# creating a log filename
logfilename="C:\\Cutover\\cchouston_log_"+datetime.now().strftime("%H%M%S")+".txt"
#creating a image logo filename
imagelogoname="c:\\cutover\\cccore1.png"

def convert_file_to_base64(filename):
    try:
        contents = open(filename, 'rb').read()
        encoded = base64.b64encode(contents)
    except Exception as error:
        print('Cancelled - An error occurred')
    return encoded

def volumebutton():
    while True:
        pyautogui.press('volumedown')
        pyautogui.press('volumedown')
        print("Volume Down")
        time.sleep(2)
        pyautogui.press('volumeup')
        print("Volume Up")
        time.sleep(120)

def callrecognize():
    recg1=recognizevoice.recognizevoice()
    recg1.speech_recognize_keyword_locally_from_microphone(logfilename)

def speech_recognize_keyword_locally_from_microphone():
    while(True):
        try:
            """runs keyword spotting locally, with direct access to the result audio"""
            speechregenabled=texttospeech.convert("Voice Recognition enabled. Please say ccHouston followed by whatever text you need.")
            speechregenabled.readout()
            x=0
            while(x<5):
                print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the time inside the function \n"))

                # Creates an instance of a keyword recognition model. Update this to
                # point to the location of your keyword recognition model.
                model = speechsdk.KeywordRecognitionModel("c:\\cutover\\ccHouston_0217.table")

                # The phrase your keyword recognition model triggers on.
                keyword = "ccHouston"

                # Create a local keyword recognizer with the default microphone device for input.
                keyword_recognizer = speechsdk.KeywordRecognizer()
                done = False

                def recognized_cb(evt):
                    # Only a keyword phrase is recognized. The result cannot be 'NoMatch'
                    # and there is no timeout. The recognizer runs until a keyword phrase
                    # is detected or recognition is canceled (by stop_recognition_async()
                    # or due to the end of an input file or stream).
                    result = evt.result
                    if result.reason == speechsdk.ResultReason.RecognizedKeyword:
                        print("RECOGNIZED KEYWORD: {}".format(result.text))
                    nonlocal done
                    done = True

                def canceled_cb(evt):
                    result = evt.result
                    if result.reason == speechsdk.ResultReason.Canceled:
                        print('CANCELED: {}'.format(result.cancellation_details.reason))
                    nonlocal done
                    done = True

                # Connect callbacks to the events fired by the keyword recognizer.
                keyword_recognizer.recognized.connect(recognized_cb)
                keyword_recognizer.canceled.connect(canceled_cb)


                # Start keyword recognition.
                result_future = keyword_recognizer.recognize_once_async(model)
                print('Say something starting with "{}" followed by whatever you want...'.format(keyword))
                result = result_future.get()

                # Read result audio (incl. the keyword).
                if result.reason == speechsdk.ResultReason.RecognizedKeyword:
                    print ("Recording started")
                    print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the record start time \n"))
                    strrecgzd=speech2text.from_mic()
                    print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the record end time \n"))
                    print(strrecgzd)


                # If active keyword recognition needs to be stopped before results, it can be done with
                #
                #   stop_future = keyword_recognizer.stop_recognition_async()
                #   print('Stopping...')
                #   stopped = stop_future.get()
                print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the time getting out \n"))
                x=x+1
            speechregenabled=texttospeech.convert("Voice Recognition ended.")
            speechregenabled.readout()
        except (NotImplementedError):
            print("Caught it.... This was missing")

def launchcommcenter():
    f = open(logfilename, "w")
    f.write("Woops! I have deleted the content!")
    f.close()
    # ending a filename

    layout0 =  [[sg.Image((convert_file_to_base64(imagelogoname)), size=(800,350), pad=(10,10))],
                [sg.HSeparator()],
                [sg.Text(size=(25,1),text='Enter Your Name:'), sg.Input(size=(25,1), key='-UserName-',background_color="cyan")],
                [sg.HSeparator()],
                [sg.Text(size=(25,1),text=''),sg.Button('Enter ccHOUSTON')]
                ]

    layout =  [[sg.Image((convert_file_to_base64(imagelogoname)), size=(800,350), pad=(10,10))],
               [sg.HSeparator()],
               [sg.Text('Cutover Options:'), sg.Text(size=(15,1), key='-OUTPUT-')],
               [sg.HSeparator()],
               [sg.Text(size=(25,1),text='Daisy Chain Tap :'),sg.Input(size=(15,1), key='-Act_Done-'),sg.Button('Check & Tap Next tasks')],
               [sg.Text(size=(15,1), key='-DaisyChainOutput-')],
               [sg.HSeparator()],
               [sg.Text(size=(25,1),text='Tap Task ID:'),sg.Input(size=(15,1), key='-TaskIDtoTap-'),sg.Button('NoCheckTap'), \
                sg.Button(button_text='Record', image_size=(45,45),image_data=convert_file_to_base64(recordingfilename),button_color=('blue','green'),border_width=0, key='Record and Tap Tasks without Checks')],
               [sg.Text(size=(15,1), key='-TapTaskIDOutput-'),sg.Text('(Eg: 34567, 24567, 21323)')],
               [sg.HSeparator()],
               [sg.Text(size=(25,1),text='Tap Tasks after time:'),sg.Input(size=(15,1), key='-TaskTime-'),sg.Button('Tap Tasks at time')],
               [sg.Text(size=(15,1), key='-TapTaskAfterTimeOutput-'),sg.Text('(Eg: 07/03/22 21:07:05)')],
               [sg.HSeparator()],
               [sg.Button('Exit')]
               ]
    window = sg.Window('ccHouston Command Center', layout0, size=[850,850])

    p2 = multiprocessing.Process(target = volumebutton, args=())
    p2.start()

    p3= multiprocessing.Process(target = callrecognize, args=())
    p3.start()

    while True:
        event, values = window.read()
        print(event, values)

        if event == sg.WIN_CLOSED: # if user closes window or clicks cancel
            if p2.is_alive():
                p2.terminate()
            if p3.is_alive():
                p3.terminate()
            break

        if event == 'Exit': # if user clicks Exit
            if p2.is_alive():
                p2.terminate()
            if p3.is_alive():
                p3.terminate()
            break

        if event == 'Enter ccHOUSTON':
            #text = "Hello, " + str(values['-UserName-']) + ". Welcome to CC Houston Command Center for Dress Rehearsal 4."
            text = "Hello, " + str(values['-UserName-'])+ ". Welcome to ccHouston."
            pid=os.getpid()
            print(str(pid) + " is the current process for the ccHouston launch window")
            convert1=texttospeech.convert(text)
            convert1.readout()
            window = sg.Window('Cutover Choices', layout, size=[850,850])

            #p1 = Process(target=launchcommcenter, args=())
            #p1.start()
            #time.sleep(3600)
            #p1.join()
            #recbuttonpressed=False
        if event == 'NoCheckTap':
            #if p2.is_alive():
            #p2.terminate()
            #if p3.is_alive():
            #p3.terminate()
            print("Done.... WTF")
            print(str(values['-TaskIDtoTap-']))
            x=str(values['-TaskIDtoTap-'])
            x=x.replace(" ","")
            x=x.replace(",","")
            x = re.findall("[0-9]{5}", x)
            print(x)
            emailtasklist=scantask.tapwithoutchecks(x,(logfilename))
            scantask.email(emailtasklist)

        if event == 'Check & Tap Next tasks':
            # Update the "output" text element
            # to be the value of "input" element
            #window['-OUTPUT-'].update("Tapping Emails")
            emailtasklist=scantask.checkandtapnexttasks(str(values['-Act_Done-']),(logfilename))
            scantask.email(emailtasklist)

        if event == 'Tap Tasks at time':
            # Update the "output" text element
            # to be the value of "input" element
            #window['-OUTPUT-'].update("Tapping Emails")
            scantask.taptasksattime(str(values['-TaskTime-']),(logfilename))
            #scantask.email(emailtasklist)

        """
        if event == 'Record and Tap Tasks without Checks':
            recbuttonpressed= not recbuttonpressed
            print("Asked to Record and Tap Tasks without Checks")
            print("Came in with value of button pressed as " + str(recbuttonpressed))

            if(recbuttonpressed):

                start1=datetime.now()
                pid = os.getpid()
                threadName = current_thread().name
                processName = current_process().name

                print((start1.strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the start time inside the main thread"))
                print(f"{pid} * {processName} * {threadName} ---> is the Main thread...\n")

                messagebox.showinfo("Record","Talk into the microphone")
                p1 = noworkthread.noWorkThread(5)
                p1.start()
                end1=datetime.now()
                print((end1.strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the end time in the main thread"))
                p1.join()
                print(p1.valretd)
                window['-TaskIDtoTap-'].update(p1.valretd)
                recbuttonpressed= not recbuttonpressed
                #print ("Recording is off")
                messagebox.showinfo("Record","Recording Stopped")
                p1.join()
                print(p1.valretd)
            else:
                print("Recording is off")
                end1=datetime.now()
                print((end1.strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the end time"))
                #speechtotext.threading_rec(2)
                print("Came in with value of button pressed as " + str(recbuttonpressed) + "... So setting button text to " + " Record")
                #window1.Element('Record and Tap Tasks without Checks').Update(('Record','Stop')[recbuttonpressed], button_color=(('blue', ('green', 'red')[recbuttonpressed])))
                #sg.Button()
            """
if __name__ == '__main__':
    p1 = multiprocessing.Process(target = launchcommcenter, args=())
    p1.start()