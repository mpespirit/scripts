import win32com.client as win32
import tkinter as tk
import re

# GUI ----
#window = tk.Tk()
#window.title("Mail Forwarding to Onsite Portal")
#window.mainloop()

#canvas = tk.Canvas(window, width=500, height=500)
#canvas.pack()
#greeting = tk.label(text="Hello world")
#greeting.pack()
#frame = tk.Frame(window, bg="blue")
#frame.place(relx=0.05, rely=0.05, relwidth = 0.9, relheight=0.9)
#window.mainloop()

# OUTLOOK ----
print("Opening Outlook...")
    
outlook = win32.Dispatch("Outlook.Application")
ns = outlook.GetNamespace("MAPI")

# Get the inbox folder and path to subfolder
inbox = ns.GetDefaultFolder('6')
subfolder = input("Please enter your subfolder name (case-senstive): ")
tracking = inbox.Folders[subfolder]

check = 'n'
while (check != 'y'):
    check = input("Please confirm ["+ subfolder +"] has been cleared since last upload (y/n): ")

to = input("To (e-mail): ") 
cc = input("CC (e-mail separated by ';' or press enter to skip): ")

#Parse each mail for ticket ID
#For each ticket id, fwd mail body
messages = tracking.Items
prefix = "<tr><td align='center'>"
mid = "</td><td align='center'>"
suffix = "</td></tr>"

flag1 = False
flag2 = False
flag3 = False

for message in messages: 
    print("Opening mail...")
    body = message.Body
    split = body.splitlines()

    cust = "" 
    ship_date = ""
    track_nums = ""
    rma_nums = ""
    ref_nums = ""
    addr = ""    
    items = []
    tickets = []
    item_tbl= ""
    
    for x in split:
        if "Hello" in x: 
            cust = x[6:]
        
        elif "Shipped Date" in x:
            ship_date = x[14:]
            
        elif "Label" in x:
            track_nums += x + "<br>"
        
        elif "RMA#:" in x:
            rma_nums = x[6:]
        
        elif "Ref #" in x:        
            ref_nums = x[13:]
        
        elif "Onsite Service Ticket" in x:
            substring = x[25:]
            alltickets = substring.split(";")
            
            #check if substring matches regex
            for id in alltickets:
                match = re.search(r'[A-Z]{3}\-\d{3}\-\d{5}', id)
                if match:
                    id = match.group()
                    tickets.append(id)
                    
        elif "Ship To" in x:
            flag1 = True
            
        elif flag1 is True:
            if "Item Details" in x:
                flag1 = False
                flag2 = True
            else: 
                addr += x + "<br>"
            
        elif flag2 is True:
            if "Note:" in x: 
                flag2 = False
              
            elif "Model" not in x and "Qty" not in x: 
                #remove leading and trailing whitespace and split values
                val = x.strip()
                if val:
                    items.append(val.split("\t"))    
        
    #turn item table into single HTML line
    for item in items:            
        for y in item:
            if flag3 is False: 
                item_tbl += prefix + y + mid 
                flag3 = True
                
            elif flag3 is True:
                item_tbl += y + suffix
                flag3 = False
    
    #remove duplicate ticket entries
    tickets = list(set(tickets))
    
    #formatted mail via HTML
    html_body = (r"""        
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"><p><style type="text/css"> 
        body,td { color:#2f2f2f; font:11px/1.35em Verdana, Arial, Helvetica, sans-serif; }
        </style><body style="background:#F6F6F6; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; margin:0; padding:0;">
       
       <div style="background:#F6F6F6; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:12px; margin:0; padding:0;">
            <table cellspacing="0" cellpadding="0" border="0" height="100%" width="100%">
                <tr>
                    <td align="center" valign="top" style="padding:20px 0 20px 0">

                    <!-- [ header starts here] -->
                    <table bgcolor="#FFFFFF" cellspacing="0" cellpadding="10" border="0" width="650" style="border:1px solid #E0E0E0;">
                        <tr>
                            <td valign="top"><img src="http://store.supermicro.com/media/email/logo/default/logo2.png" alt="Supermicro RMA" style="margin-bottom:10px;" border="0"></td>
                        </tr>
                        <!-- [ middle starts here] -->
                        <tr>
                            <td valign="top">
                                <h1 style="font-size:22px; font-weight:normal; line-height:22px; margin:0 0 11px 0;">Hello """+ cust +"""</h1>

                            <p style="font-size:12px; line-height:16px; margin:0 0 10px 0;">
                            We’re happy to let you know that one or more of your replacement items have been shipped. Please see below for details: 
                            <br><br>
                            <b>Shipped Date: </b>"""+ship_date+"""<br>
                            <b>Tracking Number: </b><br>"""+track_nums+"""
                            <b>RMA#: </b>"""+rma_nums+"""<br>
                            <b>Ref #/PO #: """+ref_nums+"""</b><br>
                            <b>Onsite Service Ticket #: </b>"""+substring+"""<br>
                            <b>Ship To:</b><br>""" + addr + """
                            </p>
                            </td>
                        </tr>
                        <tr><td><b>Item Details:</b><br>
                        <table style="border-style: solid; border-width: thin; width: 350px; border-collapse: collapse;" frame="border" ;="" border="1">
                            <tr>                        
                                <td align="center" style="font-weight: bold; background-color: #336699; color: #FFFFFF;" width="200"> &nbsp; Model</td>
                                <td align="center" style="font-weight: bold; background-color: #336699; color: #FFFFFF;" width="150"> &nbsp; Qty</td>
                            </tr> 
                        """+ item_tbl + """
                        </table>                    
                        <tr>
                            <td>
                                <p>Note: This email was sent from a notification-only email address that cannot accept incoming e-mail. Please do not reply to this message. Please kindly contact us at rma@supermicro.com
                                </p>
                            </td>
                        </tr>
                    </table>
                    </td>
                </tr>
            </table>
        </div></body><p>
    """)

    #post RMA details to tickets in this RMA
    for ticket in tickets:  
        #create email to send to onsite portal
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = "[#"+ ticket +"]" 
        mail.HTMLBody = html_body
        
        try:         
            print("Sending Mail to Ticket ID: "+ticket)
            mail.Send()
        except:
            print("Cannot send mail")
    
    flag1 = False
    flag2 = False
    flag3 = False
    print("Closing mail...")