import win32com.client as win32
import sys

outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)#.Folders.Item("Your_Folder_Name")

def extract_file(path,send_name,send_date,file_name,output_name):
    emails = []
    for message in inbox.Items:
        this_message = (
        message.Subject,
        message.Attachments,  
        message.ReceivedTime.strftime("%Y-%m-%d"),
        message.Sender.Address)
        emails.append(this_message)

        for email in emails:
            subject,attachments,dt,sender = email 

            if dt==send_date and sender==send_name: # change the date and sender name

                for attach in attachments:
                    if attach.FileName == file_name:
                        attach.SaveAsFile(path +'\\' + output_name+'.xlsx')# + attach.FileName)        

def read_excel_file(output_name,sheet_name):
    df = pd.read_excel(output_name,sheet_name =sheet_name) # if no sheet name then write FALSE
    return df

def read_csv_file(output_name):
    df = pd.read_csv(output_name)
    return df
    
def main():
    extract_file(r"C:\Users\",'@emailaddress','2022-03-31','outlook.xlsx','2022-04-02')
    

if __name__ == "__main__":
    sys.exit(main())     