import math
import time
import pandas as pd
from datetime import timedelta, datetime
import scantask
from readfile import Task
from itertools import groupby
import sendmail


def checkandtapnexttasks(taskid,filename):

    f = open(filename, "a")
    print("Came in with Tapping task : " + str(taskid))
    f.write("\n"+ " Doing a daisy chain scan for task "+ str(taskid))
    #df=pd.read_excel("C:\\Users\\pradhip.swaminathan\\IdeaProjects\\Exceltomail\\cutover_sample.xlsx")
    df=pd.read_excel("C:\\Cutover\\cutover_sample.xlsx")
    taskstotap=[]
    tasklist=[]
    tasktoevaluate=[]

    for x in range (0,(df.shape[0])):
        tasklist.append(Task(x, \
                             (df['Start Mode'].loc[df.index[x]]), \
                             (df['Activity ID'].loc[df.index[x]]), \
                             (df['Predecessors'].loc[df.index[x]]), \
                             (df['Successors'].loc[df.index[x]]), \
                             (df['Task Name'].loc[df.index[x]]), \
                             (df['Duration'].loc[df.index[x]]), \
                             (str(df['Start Time'] .loc[df.index[x]])),
                             (df['Primary'].loc[df.index[x]]),
                             (df['Secondary'].loc[df.index[x]]),
                             (str(df['Finish Time'].loc[df.index[x]])),
                             (str(df['Completion'].loc[df.index[x]])),
                             (str(df['Workstream'].loc[df.index[x]])),
                             (str(df['Simple Task Name'].loc[df.index[x]]))
                             ))
        if(str(tasklist[x].ActivityID)==str(taskid)):
            tasktoevaluate.append(tasklist[x])
            print("Added original Task " + str(taskid) + " to be evaluated to evaluate list :")
            for i in range (0,len(tasktoevaluate)):
                print (str(tasktoevaluate[i].ActivityID))


    #for loopctr in tasktoevaluate:
    try:
        for loopctr in range (0,16):
            print("Now evaluating " + str(taskid))
            for x in range (0,(df.shape[0])):
                #print(str(tasklist[x].ActivityID)+ " --- " + tasklist[x].CompletionPercent)
                if((str(tasklist[x].ActivityID)==str(taskid)) and ((float(tasklist[x].CompletionPercent))!= 25.0)):
                    taskexistsflag=0
                    for ctr in range (0, len(tasktoevaluate)):
                        if (str(tasktoevaluate[ctr].ActivityID)==str(taskid)):
                            taskexistsflag=1
                            print ("Task " + str(taskid) + " already in evaluate list. No need to add to evaluate list again")
                            if (taskexistsflag==0):
                                tasktoevaluate.append(tasklist[x])
                                print("Task ID  " + str(taskid) + " added to evaluation list. Being set to completion of 25")
                                taskstotap[len(taskstotap)-1].CompletionPercent="25"
                                tasklist[x].CompletionPercent="25"
                    #tasktoevaluate.append(tasklist[x])
                    #print("Added Task " + str(taskid) + " to be evaluated to evaluate list :")

                    #for i in range (0,len(tasktoevaluate)):
                    #    print (str(tasktoevaluate[i].ActivityID))

                    for z2 in (str(tasklist[x].Predecr).split(",")):
                        if( not(z2) or ((len(z2))>5)):
                            print (str(taskid) + " has no predecessors or has incorrect formatting of predecessors")
                            f.write("\n"+ str(taskid) + " has no predecessors or has incorrect formatting of predecessors")
                        else:
                            print((z2) + " are the predecessors for " + str(taskid))
                            f.write("\n"+ (z2) + " are the predecessors for " + str(taskid))
                            print(str(len(taskstotap)) + " are the number of tasks to tap")
                            for z3 in range (0,len(tasklist)):
                                taskexistsintappinglistflag=0
                                if( (str(tasklist[z3].ActivityID)==z2)):
                                    print("Checking predecessor completion " + z2)
                                    for ctr in range (0, len(taskstotap)):
                                        print(str(taskstotap[ctr].ActivityID) + " ~~~ " + str(z2))

                                        if (str(taskstotap[ctr].ActivityID)==str(z2)):
                                            print("Now checking if " + z2 + " is in tapping list")
                                            taskexistsintappinglistflag=1
                                            print ("Predecessor Task " + str(z2) + " already in tapping list. No need to check for it's completion")
                                    if (taskexistsintappinglistflag==0):
                                        z7=float(tasklist[z3].CompletionPercent)
                                        z7=z7*100
                                        print (str(tasklist[z3].ActivityID) + " has " + str(z7) + " as completion..")
                                        if ( ((z7 < 90.0)  or (math.isnan(z7))) and (z7 != 25.0) ):
                                            print("Task " + str(taskid) + " has a Predecessor  " + z2 + " which is not complete")
                                            f.write("\n"+ "Task " + str(taskid) + " has a Predecessor  " + z2 + " which is not complete")
                                            raise StopIteration


                    print("Task ID  " + str(taskid) + " looks good for tapping e-mail. All predecessors good.")
                    f.write("\n"+ "Task ID  " + str(taskid) + " looks good for tapping e-mail. All predecessors good.")
                    if (len(taskstotap)==0):
                        print (float(tasklist[x].CompletionPercent))
                        if((float(tasklist[x].CompletionPercent) * 100.0) > 0.0):
                            print("Task ID " + str(tasklist[x].ActivityID)+ " has been tapped already. Adding now will be removing soon")
                            f.write("\n"+"Task ID " + str(tasklist[x].ActivityID)+ " has been tapped already. Adding now will be removing soon")
                            taskstotap.append(tasklist[x])
                        else:
                            taskstotap.append(tasklist[x])
                            print("Tapping list is " + str(len(taskstotap)) )
                            print("Task ID  " + str(taskid) + " added to tapping list. Being set to completion of 25")
                            taskstotap[len(taskstotap)-1].CompletionPercent="25"
                            tasklist[x].CompletionPercent="25"
                    else:
                        taskexistsflag=0
                        print("I am here... Checking to see if task " + str(taskid) + " needs to be added to tapping list")
                        for ctr in range (0, len(taskstotap)):
                            #print ("I find task " + str(taskstotap[ctr].ActivityID) + " in tapping list....")
                            if (str(taskstotap[ctr].ActivityID)==str(taskid)):
                                taskexistsflag=1
                                print ("Task " + str(taskid) + " already in tapping list. No need to add it again")
                            for ctr in range (0,(df.shape[0])):
                                if((str(tasklist[ctr].ActivityID)==str(taskid)) and (((float(tasklist[ctr].CompletionPercent))*100)>= 20.0) ):
                                    #print (tasklist[ctr].CompletionPercent)
                                    taskexistsflag=1
                                    print ("Task " + str(taskid) + " has already been tapped. No need to add it again")
                        if (taskexistsflag==0):
                            taskstotap.append(tasklist[x])
                            print("Tapping list is " + str(len(taskstotap)) )
                            print("Task ID  " + str(taskid) + " added to tapping list. Being set to completion of 25")
                            taskstotap[len(taskstotap)-1].CompletionPercent="25"
                            tasklist[x].CompletionPercent="25"


                    print("Now evaluating successors for task ..." + str(tasklist[x].ActivityID))
                    z7=str(tasklist[x].Succesr)
                    z7=z7.replace(", ",",")
                    print(z7)
                    for z in (z7.split(",")):
                        if(len(z)>5):
                            print("Successors for task " + str(tasklist[x].ActivityID) + " has incorrect formatting")
                        for y in range (0, len(tasklist)):
                            z5=str(tasklist[y].ActivityID)
                            #print (z + " being compared with " + z5 )
                            if ((z5==z)):
                                if(((float(tasklist[y].CompletionPercent))*100)!=100.0):
                                    print(str(tasklist[y].ActivityID) + " successor task is added to list for tapping evaluation.")
                                    tasktoevaluate.append(tasklist[y])
                                    print("Added Task " + str(tasklist[y].ActivityID) + " to be evaluated to evaluate list:")
                                else:
                                    print(str(tasklist[y].ActivityID) + " is a successor task which is already done...")
                        #for i in range (0,len(tasktoevaluate)):
                        #print(str(tasktoevaluate[i].ActivityID) + " --- " + (str(tasktoevaluate[i].CompletionPercent)))

            #print("Removing task " + str(tasktoevaluate[0].ActivityID) + " from Activity List...")
            tasktoevaluate.pop(0)
            #if(loopctr==0):
            #    print("Removing task " + str(tasktoevaluate[0].ActivityID) + " from Activity List...")
            #    tasktoevaluate.pop(0)
            if(len(tasktoevaluate)>0):
                taskid=tasktoevaluate[0].ActivityID
            else:
                raise StopIteration

        for ctr in range (0, len(taskstotap)):
            ctr1=0
            #print (str(ctr1))
            #print((str(taskstotap[ctr1].ActivityID)) + " ---- " + taskstotap[ctr1].TaskName + " ---- " + (str(taskstotap[ctr1].CompletionPercent)))
            z=(float(taskstotap[ctr1].CompletionPercent))
            if ((float(taskstotap[ctr1].CompletionPercent))>0.0) and (((float(taskstotap[ctr1].CompletionPercent))<0.26)):
                #print("Looks like a tapped task@@@")
                taskstotap.pop(ctr1)
                ctr1=ctr1-1
                #print(str(ctr1))
        return(taskstotap)
    except StopIteration:
        for ctr in range (0, len(taskstotap)):
            ctr1=0
            #print (str(ctr1))
            #print((str(taskstotap[ctr1].ActivityID)) + " ---- " + taskstotap[ctr1].TaskName + " ---- " + (str(taskstotap[ctr1].CompletionPercent)))
            z=(float(taskstotap[ctr1].CompletionPercent))
            if ((float(taskstotap[ctr1].CompletionPercent))>0.0) and (((float(taskstotap[ctr1].CompletionPercent))<0.26)):
                #print("Looks like a tapped task@@@")
                taskstotap.pop(ctr1)
                ctr1=ctr1-1
                #print(str(ctr1))
        f.close()
        return(taskstotap)

def taptasksattime(timetocheck,filename):
    f = open(filename, "a")
    f.write("\n"+ " Doing a scan for tasks around "+ str(timetocheck))
    print("Hello there... Scanning for tasks around " + timetocheck)
    checktimevariable=datetime.strptime(str(timetocheck),"%m/%d/%y %H:%M:%S")
    print(checktimevariable)
    checkstarttime=(checktimevariable)-timedelta(minutes=600)
    checkendtime=(checktimevariable)+timedelta(minutes=30)
    f.write("\n"+"Checking tasks with starttime between " +  str(checkstarttime) + " and " + str(checkendtime))
    #df=pd.read_excel("C:\\Users\\pradhip.swaminathan\\IdeaProjects\\Exceltomail\\cutover_sample.xlsx")
    df=pd.read_excel("C:\\Cutover\\cutover_sample.xlsx")
    tasklist=[]
    for x in range (0,(df.shape[0])):
        tasklist.append(Task(x, \
                             (df['Start Mode'].loc[df.index[x]]), \
                             (df['Activity ID'].loc[df.index[x]]), \
                             (df['Predecessors'].loc[df.index[x]]), \
                             (df['Successors'].loc[df.index[x]]), \
                             (df['Task Name'].loc[df.index[x]]), \
                             (df['Duration'].loc[df.index[x]]), \
                             (str(df['Start Time'] .loc[df.index[x]])),
                             (df['Primary'].loc[df.index[x]]),
                             (df['Secondary'].loc[df.index[x]]),
                             (str(df['Finish Time'].loc[df.index[x]])),
                             (str(df['Completion'].loc[df.index[x]])),
                             (str(df['Workstream'].loc[df.index[x]])),
                             (str(df['Simple Task Name'].loc[df.index[x]]))
                             ))

    ctr=0
    taskstotap=[]

    for x in range (0,(df.shape[0])):
        print ("-----------------------------------------------------------------------------------------------------------------")
        print(str(tasklist[x].ActivityID) + " --- " + tasklist[x].CompletionPercent )
        print(type(tasklist[x].CompletionPercent))
        tappercent=(float(tasklist[x].CompletionPercent))
        print(tappercent==0.0)
        print(math.isnan(tappercent))
        print (str(x) + " . "+ " Evaluating " + str(tasklist[x].ActivityID) + " ~~~" + str(tappercent)+"% " + "~~~" + tasklist[x].StartTime)
        taskstarttime=(datetime.strptime(tasklist[x].StartTime,"%Y-%m-%d %H:%M:%S"))
        checkstarttime=(checktimevariable)-timedelta(minutes=300)
        checkendtime=(checktimevariable)+timedelta(minutes=30)
        print("Checking tasks with starttime between " +  str(checkstarttime) + " and " + str(checkendtime))

        if(ctr<100):
            if ((taskstarttime>=checkstarttime) and (taskstarttime<=checkendtime)  and (math.isnan(tappercent) or (tappercent==0.0)) ):
                print(' Going to check daisy chain links and tap next tasks for :' + str(tasklist[x].ActivityID) + " Round " + str(ctr+1))
                f.write("\n"+' Going to check daisy chain links and tap next tasks for :' + str(tasklist[x].ActivityID) + " Round " + str(ctr+1))
                taskstotap=checkandtapnexttasks(tasklist[x].ActivityID,filename)
                scantask.email(taskstotap)
                ctr=ctr+1
    f.close()
    return taskstotap
    #print ("Please enter the input date in format mm/dd/yy hh:mi:ss (Eg: 07/03/22 21:07:05) ")

def tapwithoutchecks(tasklist,filename):

    f = open(filename,"a")
    #print("Checking for tasks(s) " + tasklist + " for tapping without checks")
    #f.write("\n"+ "Checking for tasks(s) " + tasklist + " for tapping without checks")
    #df=pd.read_excel("C:\\Users\\pradhip.swaminathan\\IdeaProjects\\Exceltomail\\cutover_sample.xlsx")
    df=pd.read_excel("C:\\Cutover\\cutover_sample.xlsx")
    tasktoemail=[]
    successmessage=""
    #for y in (str(tasklist).replace(", ",",").split(",")):
    for y in tasklist:
        print(str(y) + " being checked.")
        f.write("\n"+str(y) + " being checked.")
        taskfindctr=0
        for x in range (0,(df.shape[0])):
            #print("Comparing " + str(x) + " with " + str(y))
            if( (str(y)==str(df['Activity ID'].loc[df.index[x]]))):
                print("Found the task for tapping")
                #print(df['Completion'].loc[df.index[x]])
                if((df['Completion'].loc[df.index[x]]>0.01)):
                    print(("Found a match, but this task ") + (str(y)) + " has already been tapped....")
                    f.write("\n"+"Found a match, but this task " + (str(y)) + " has already been tapped....")
                    taskfindctr=1
                if((df['Completion'].loc[df.index[x]]<0.01) or math.isnan(df['Completion'].loc[df.index[x]])):
                    print(("Found a match and this task ") + (str(y)) + "  and it's going to be tapped....")
                    f.write("\n"+"Creating Tapping e-mail for task " + (str(y)) + "....")
                    successmessage=successmessage + "Task " + str(y) + ", "
                    print(successmessage)
                    taskfindctr=1
                    tasktoemail.append(Task(x, \
                                            (df['Start Mode'].loc[df.index[x]]), \
                                            (df['Activity ID'].loc[df.index[x]]), \
                                            (df['Predecessors'].loc[df.index[x]]), \
                                            (df['Successors'].loc[df.index[x]]), \
                                            (df['Task Name'].loc[df.index[x]]), \
                                            (df['Duration'].loc[df.index[x]]), \
                                            (str(df['Start Time'] .loc[df.index[x]])),
                                            (df['Primary'].loc[df.index[x]]),
                                            (df['Secondary'].loc[df.index[x]]),
                                            (str(df['Finish Time'].loc[df.index[x]])),
                                            (str(df['Completion'].loc[df.index[x]])),
                                            (str(df['Workstream'].loc[df.index[x]])),
                                            (str(df['Simple Task Name'].loc[df.index[x]]))
                                            ))
        if (taskfindctr ==0):
            print("Task " + str(y)+ " could not be found in the list for tapping.")
            f.write("\n"+"Task " + str(y)+ " could not be found in the list for tapping.")
    #print(str(len(tasktoemail)) + " is the length")
    #print (successmessage)
    print("Out of the loop.")
    print(successmessage)
    if (len(successmessage)>0):
        time.sleep(1)
        #scsmsg=texttospeech.convert(successmessage + " are ready to be e-mailed")
        #scsmsg.readout()
    f.close()
    return(tasktoemail)

def email(tasklist):
    print("preparing to email " + str(len(tasklist)) + " number of tasks...")
    tasklist.sort(key=lambda x: x.Workstream)
    groupedwks = [list(result) for key, result in groupby(
        tasklist, key=lambda x: x.Workstream)]
    rows=len(groupedwks)
    for ctr in range (0,rows):
        starttime=str(groupedwks[ctr][0].StartTime)
        html_body = """
                <!-- HTML Codes by Quackit.com -->
                <!DOCTYPE html>
                <title>Text Example</title>
                <style>
                div.container {
                background-color: #ffffff;
                }
                div.container p {
                font-family: Arial;
                font-size: 14px;
                font-style: normal;
                font-weight: normal;
                text-decoration: none;
                text-transform: none;
                color: #000000;
                }
                .tg  {border-collapse:collapse;border-spacing:0;}
                .tg td{border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;
                  overflow:hidden;padding:10px 5px;word-break:normal;}
                .tg th{border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;
                  font-weight:normal;overflow:hidden;padding:10px 5px;word-break:normal;}
                .tg .tg-7dil{background-color:#009901;text-align:left;vertical-align:top}
                .tg .tg-0lax{text-align:left;vertical-align:top}
                </style>
                <div class="container">
                <p>Please find below your upcoming tasks for <span style="color:blue">%s</span>. (This is when this email should have been sent, but still please send the real actual start and finish times);</p>
                <p></p>
                <p>•    Please reply back to the tapping e-mail or call into the Rate Refund 2 Command Center Open Line (Log into Teams Meeting and use VTC: 116 193 742 3) to acknowledge that you have received the tapping email and that you are beginning your tasks.</p>
                <p>•	When you have completed your tasks, please respond back to the tapping email with the Actual Start Time and Actual End Time <U>in EST</U>. </p>
                <p>•	In addition, please email back or dial back into the CoreCutover Command Center Open Line to confirm task progress if it's a long running task (more than one hr).</p>
                <p>We are requesting folks to please provide the Actual Start Time and Actual End Time to the task for us to capture actual duration for future planning.</p>
                <p> </p>
                <p>Once you have completed the task, please respond with the filled out table below:</p>
                <p> </p>
                <p>@CC_PRJ_BASIS : Team pls monitor the system while conversion loads are being executed. </p>
                <p></p>
                <p></p>
                <p></p>
                <p></p>
                <p></p>
                </div>
            
            <table class="tg">
            <thead>
              <tr>
                <th class="tg-7dil">Unique <br>ID</th>
                <th class="tg-7dil">Percent<br>Comp</th>
                <th class="tg-7dil">Task Name</th>
                <th class="tg-7dil">Workstream</th>
                <th class="tg-7dil">Primary Owner</th>
                <th class="tg-7dil">Secondary Owner</th>
                <th class="tg-7dil">Planned <br>Duration</th>
                <th class="tg-7dil">Planned <br>Start</th>
                <th class="tg-7dil">Planned <br>Finish</th>
                <th class="tg-7dil">Actual Start</th>
                <th class="tg-7dil">Actual Finish</th>
              </tr>
            </thead>
            <tbody>
        """% (starttime)
        html_body2=""
        activityid_subj=""
        activityid_primary=""
        activityid_cc=""
        for ctr1 in range(0,len(groupedwks[ctr])):
            varStartTime=datetime.strptime(groupedwks[ctr][ctr1].StartTime,"%Y-%m-%d %H:%M:%S")
            varStartTime=varStartTime.strftime("%m/%d/%y %H:%M")
            varEndTime=datetime.strptime(groupedwks[ctr][ctr1].EndTime,"%Y-%m-%d %H:%M:%S")
            varEndTime=varEndTime.strftime("%m/%d/%y %H:%M")
            activityid_subj=activityid_subj+str(groupedwks[ctr][ctr1].ActivityID)+"~~~"
            activityid_primary=activityid_primary+str(groupedwks[ctr][ctr1].Primary)+"; "
            activityid_cc=activityid_cc+str(groupedwks[ctr][ctr1].Secondary)+"; "
            print(str(groupedwks[ctr][ctr1].ActivityID) + "----" + groupedwks[ctr][ctr1].Workstream + "----" + groupedwks[ctr][ctr1].TaskName)
            html_body1="""
                    <tr>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">%s</td>
                        <td class="tg-0lax">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        <td class="tg-0lax">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                    </tr>"""%(groupedwks[ctr][ctr1].ActivityID,groupedwks[ctr][ctr1].CompletionPercent, \
                              groupedwks[ctr][ctr1].TaskName, \
                              groupedwks[ctr][ctr1].Workstream, \
                              groupedwks[ctr][ctr1].Primary,groupedwks[ctr][ctr1].Secondary, \
                              groupedwks[ctr][ctr1].Duration, \
                              varStartTime, \
                              varEndTime)
            html_body2=html_body2+html_body1
        html_body3="""
        </tbody>
        </table>"""
        sendmail.create_email(activityid_primary,activityid_cc,(str("[ACTION REQUESTED] - Customer Core Rate Refund 2 Cutover - (") + str(activityid_subj) + ")"),(html_body+html_body2+html_body3))
