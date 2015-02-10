import win32com.client
import os, time, sys
import re
import subprocess

Manager = 'Your Managers Domain Name'



def set_color(color):
    colorCmd = '--rgb='+str(color[0])+','+str(color[1])+','+str(color[2])
    cmd =['blink1-tool','-q',colorCmd]
    pr=subprocess.Popen(cmd)
    pr.wait()


print 'I will keep an eye out for your emails :) take a walk..'
session = win32com.client.gencache.EnsureDispatch ("MAPI.Session")

session.Logon ("Outlook")
messages = session.Inbox.Messages


#while 1:
while (1):
    message = messages.GetLast ()
    if (message.Unread):
        sender = str(message.Sender.Address)
        #print sender
        matchObj = re.match( r'.*/cn=(.*)', sender, re.M|re.I)
        try:
            From = matchObj.group(1)
            ##print From ## uncomment this to get the address 
					   ## header you might have to change 
					   ##the regex to get this working
        except:
            From = 'unknown'
			
			
            
        if (From == Manager):
            ##print 'pattern:emergency'
            ##print '#AB0005'
            set_color([130,0,0])
            while(1):
                message = messages.GetLast ()
                if (message.Unread == False):
                    set_color([0,130,0])
                    break
                time.sleep(2)
            #print 'you are screwed'
        else:
            #print 'pattern:yellow flashes'
            #print '#AB9700'
            #print 'you have new emails'
            set_color([0,0,130])
        
    else :
        #print 'pattern:green flashes'
        #print '#00AB19'
        set_color([0,130,0])
        #print 'you have no new email'
        ##message = messages.GetNext ()
    time.sleep (60)