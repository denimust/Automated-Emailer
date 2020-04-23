import pandas
from nameparser import HumanName
from tkinter import filedialog
from tkinter import *
import easygui
root = Tk()
root.filename = filedialog.askopenfilename(initialdir = "default path",title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))

#Opening the data set and making lists from the columns that will be used.
data= pandas.read_csv(str(root.filename))

wo = list(data["WORKORDER_ID"])

type = list(data["SUBTYPE"])

email=list(data["REPORTED_BY_EMAIL"])

name=list(data["REPORTED_BY_NAME"])

building=list(data["BUILDING_NAME"])

cc=list(data["SALESPERSON_NAME"])

cc2=list(data["SALESMANAGER_NAME"])

alt_wo=list(data["ALTNUMBER"])

email2=list(data["QUESTIONSTO_NAME"])

charge_type = list(data["INVOICETYPE"])




#The function that I use to have Outlook open a new email window with the needed information.
def Emailer(text1, subject, recipient, recipient2, cc, bcc):
    import win32com.client as win32

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient + recipient2
    mail.Cc = cc
    mail.Bcc = bcc
    mail.Subject = subject
    mail.HtmlBody = text1
    mail.Display(False)

#Variable for yesterday's date or completion date. I modified it to enter the date myself as date of completion wasn't always the day
#prior
yesterday_date= easygui.enterbox("Enter date of completion")








#Loop for the email that will be made. Body of text can be altered here. chrg_tp determines if the work was warranty related or not.
for chrg_tp,w, tp, ml, nm, bldg, c, c2, alt_w, ml2 in zip(charge_type,wo, type, email, name,building, cc, cc2, alt_wo, email2):
    if chrg_tp == "EXTERNAL CHARGE":
        Emailer("""<p>Good Morning {0} and {1}!
                </p><p>Our service team completed a work order at {2} (WO# {4}) on {3}.
                </p><p>Any necessary closing documentation will follow soon from our A.R. department.
                </p><p>Please let us know if you have any questions or need any additional information regarding the completed work.
                </p><p>All of us at Company thank you for your continued business!</p>""".format(
                HumanName(nm).first,
                HumanName(ml2).first,
                bldg,
                yesterday_date,
                str(alt_w)).replace(" and !", "!", 1).replace("T&M Roof Leak", "leak repairs", 1).replace("T&M Waterproofing Leak", "leak repairs", 1).replace("new penetration installation", "a new penetration installation", 1).replace("RECALL 2nd Trip", "", 1).replace("RECALL 3rd+ Trip", "", 1).replace("Roof Repair Work", "roof repairs", 1).replace("(from KPC report)", "", 1).replace("T&M", "a", 1).replace("combined", "", 1).replace(
                "Sealant Work", "sealant work", 1).replace("Roof Assessment", "a roof assessment", 1).replace("(WO# nan)", "", 1).replace(".0", "", 1),

                "Completed WO {0} | {1}".format(w,
                bldg),

                "{0}".format(ml).replace("nan", "No Email Provided", 1),
                ";{0}".format(ml2), c, c2)
    else:
        Emailer("""<p>Good Afternoon {0} and {1}!
                        </p><p>Our service team completed a work order at {2} (WO# {4}) on {3}.
                        </p><p>Any necessary closing documentation will follow soon from our A.R. department.
                        </p><p>Please let us know if you have any questions or need any additional information regarding the completed work.
                        </p><p>All of us at KPost thank you for your continued business!</p>""".format(
            HumanName(nm).first,
            HumanName(ml2).first,
            bldg,
            yesterday_date,
            str(alt_w)).replace(" and !", "!", 1).replace("T&M Roof Leak", "leak repairs", 1).replace(
            "T&M Waterproofing Leak", "leak repairs", 1).replace("new penetration installation",
                                                                 "a new penetration installation", 1).replace(
            "RECALL 2nd Trip", "", 1).replace("RECALL 3rd+ Trip", "", 1).replace("Roof Repair Work", "roof repairs",
                                                                                 1).replace("(from KPC report)", "",
                                                                                            1).replace("T&M", "a",
                                                                                                       1).replace(
            "combined", "", 1).replace(
            "Sealant Work", "sealant work", 1).replace("Roof Assessment", "a roof assessment", 1).replace("(WO# nan)",
                                                                                                          "",
                                                                                                          1).replace(
            ".0", "", 1),

                "Completed WO {0} | {1}".format(w,
                                                bldg),

                "{0}".format(ml).replace("nan", "No Email Provided", 1),
                ";{0}".format(ml2),

                c,

                c2)
