import win32com.client as client 
outlook = client.Dispatch("Outlook.Application")
message = outlook.createItem(0)
message.display()
message.To = 'Ac933@365e.live'
message.CC = 'Ac933@365e.live'
message.Subject = "Regarding your IT incident"
message.Body = "Hi, please contact me regarding _______ when you're available, We can plan some time in to discuss"
message.Send()

