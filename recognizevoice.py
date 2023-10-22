import azure.cognitiveservices.speech as speechsdk
import scantask
import speech2text
import texttospeech
from datetime import datetime
class recognizevoice():
    # override the constructor of thread/process
    def __init__(self):
        # first execute the base constructor
        #Thread.__init__(self)
        # now store the additional value you need in this class
        self.secdelay = 5
        self.strrecgzd= ""

    #override the run function of thread/process
    def run(self):
        self.speech_recognize_keyword_locally_from_microphone()

    def speech_recognize_keyword_locally_from_microphone(self,logfilename):
        while(True):
            try:
                """runs keyword spotting locally, with direct access to the result audio"""
                #speechregenabled=texttospeech.convert("Voice Recognition enabled. Please say Hey Houston followed by whatever text you need.")
                speechregenabled=texttospeech.convert("Voice Recognition enabled")
                speechregenabled.readout()
                x=0
                while(True):
                    #while(x<10):
                    print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the time inside the function \n"))

                    # Creates an instance of a keyword recognition model. Update this to
                    # point to the location of your keyword recognition model.
                    #model = speechsdk.KeywordRecognitionModel("c:\\cutover\\ccHouston_0217.table")
                    model=speechsdk.KeywordRecognitionModel("C:\\cutover\HeyHouston_02212023.table")
                    # The phrase your keyword recognition model triggers on.
                    keyword = "Hey Houston"

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
                        self.strrecgzd=speech2text.from_mic()
                        scantask.email(scantask.tapwithoutchecks(self.strrecgzd, logfilename))
                        print((datetime.now().strftime("%m/%d/%Y, %H:%M:%S.%f"))+(" is the record end time \n"))
                        print(self.strrecgzd)


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



