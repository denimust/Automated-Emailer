import pandas
from nameparser import HumanName
from tkinter import filedialog
from tkinter import *
import easygui
root = Tk()
root.filename = filedialog.askopenfilename(initialdir = "C:\\Users\\deniz.mustafa\\Downloads",title = "Select file",filetypes = (("csv files","*.csv"),("all files","*.*")))

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

yesterday_date= easygui.enterbox("Enter date of completion")









for chrg_tp,w, tp, ml, nm, bldg, c, c2, alt_w, ml2 in zip(charge_type,wo, type, email, name,building, cc, cc2, alt_wo, email2):
    if chrg_tp == "EXTERNAL CHARGE":
        Emailer("""<p>Good Morning {0} and {1}!
                </p><p>Our service team completed a work order at {2} (WO# {4}) on {3}.
                </p><p>Any necessary closing documentation will follow soon from our A.R. department.
                </p><p>Please let us know if you have any questions or need any additional information regarding the completed work.
                </p><p>All of us at KPost thank you for your continued business!</p>""".format(
                HumanName(nm).first,
                HumanName(ml2).first,
                bldg,
                yesterday_date,
                str(alt_w)).replace(" and !", "!", 1).replace("T&M Roof Leak", "leak repairs", 1).replace("T&M Waterproofing Leak", "leak repairs", 1).replace("new penetration installation", "a new penetration installation", 1).replace("RECALL 2nd Trip", "", 1).replace("RECALL 3rd+ Trip", "", 1).replace("Roof Repair Work", "roof repairs", 1).replace("(from KPC report)", "", 1).replace("T&M", "a", 1).replace("combined", "", 1).replace(
                "Sealant Work", "sealant work", 1).replace("Roof Assessment", "a roof assessment", 1).replace("(WO# nan)", "", 1).replace(".0", "", 1),

                "Completed WO {0} | {1}".format(w,
                bldg),

                "{0}".format(ml).replace("nan", "No Email Provided", 1),
                ";{0}".format(ml2),


                "service@kpostcompany.com; {0}; {1}".format(c.replace("Doug Helixon", "doug.helixon@kpostcompany.com", 1).replace("Tracey Donels", "", 1).replace("Jill Melancon", "", 1).replace("Scott Bredehoeft", "Scott.Bredehoeft@kpostcompany.com", 1).replace("Shawn Morgan", "shawn.morgan@kpostcompany.com", 1).replace("Austin Fennema", "Austin.Fennema@kpostcompany.com", 1),
                c2.replace("Tanner Newton", "tanner.newton@kpostcompany.com", 1).replace("Adrian Hilton", "adrian.hilton@kpostcompany.com", 1).replace("Zachary Silvey", "zach.silvey@kpostcompany.com").replace("Boris Marquez", "boris.marquez@kpostcompany.com", 1).replace("Tyler Lynch", "tyler.lynch@kpostcompany.com", 1).replace("Justin Gray", "justin.gray@kpostcompany.com", 1).replace("Tracey Donels", "", 1)),

                "lori.suits@kpostcompany.com; lindsey.clark@kpostcompany.com; kristin.morris@kpostcompany.com"
        )
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

                "service@kpostcompany.com; {0}; {1}".format(
                    c.replace("Doug Helixon", "doug.helixon@kpostcompany.com", 1).replace("Tracey Donels", "",
                                                                                          1).replace("Jill Melancon",
                                                                                                     "", 1).replace(
                        "Scott Bredehoeft", "Scott.Bredehoeft@kpostcompany.com", 1).replace("Shawn Morgan",
                                                                                            "shawn.morgan@kpostcompany.com",
                                                                                            1).replace("Austin Fennema",
                                                                                                       "Austin.Fennema@kpostcompany.com",
                                                                                                       1),
                    c2.replace("Tanner Newton", "tanner.newton@kpostcompany.com", 1).replace("Adrian Hilton",
                                                                                             "adrian.hilton@kpostcompany.com",
                                                                                             1).replace(
                        "Zachary Silvey", "zach.silvey@kpostcompany.com").replace("Boris Marquez",
                                                                                  "boris.marquez@kpostcompany.com",
                                                                                  1).replace("Tyler Lynch",
                                                                                             "tyler.lynch@kpostcompany.com",
                                                                                             1).replace("Justin Gray",
                                                                                                        "justin.gray@kpostcompany.com",
                                                                                                        1).replace(
                        "Tracey Donels", "", 1)),

                "lori.suits@kpostcompany.com; lindsey.clark@kpostcompany.com; kristin.morris@kpostcompany.com")