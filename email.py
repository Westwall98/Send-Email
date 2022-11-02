from openpyxl import load_workbook
import win32com.client

def findlist():
    wb = load_workbook('C:\\Users\\Ethan\\Desktop\\Project\\user list.xlsx')
    ws = wb.active
    rows_data = list(ws.rows)
    titles = [title.value for title in rows_data.pop(0)]
    all_row_dict = []
    for a_row in rows_data:
        the_row_data = [cell.value for cell in a_row]
        row_dict = dict(zip(titles, the_row_data))
        all_row_dict.append(row_dict)

    return all_row_dict


def outlook(count):
    userlist = findlist()
    outlook = win32com.client.Dispatch("Outlook.Application")
    # outlook = win32com.client.DispatchEx("Outlook.Application")
    # outlook.Visible = 0
    msg = outlook.GetNamespace("MAPI").OpenSharedItem(r"C:\\Users\\Ethan\\Desktop\\Project\\Check-in 2022 - Basic.msg")
    for i in range(count):
        firstname = userlist[i]['User First Name']
        useremailaddr = userlist[i]['Email']
        msg.HTMLBody = msg.HTMLBody.replace('%%FirstName%%','{},'.format(firstname))
        mail = outlook.CreateItem(0)
        mail.HTMLBody = msg.HTMLBody
        mail.To = useremailaddr
        mail.Subject = '有关IT的支持，可以联系我们！'
        mail.BodyFormat = 2  # 2表示使用Html format
        # mail.Attachments.Add('C:\Users\xxx\Desktop\git_auto_pull_new.py')

        mail.Display()
        # mail.Send()


outlook(1)
