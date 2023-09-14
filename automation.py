import pandas as pd
from IPython.display import display
import numpy as np
import os
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import vobject
import shutil

from numpy.lib.twodim_base import mask_indices

master_path = '/content/drive/My Drive/SMC_scripts/smc_contact_table/'

#get earliest date
def flatten_list(list_of_lists):
    flat_list = []

    for item in list_of_lists:
        if isinstance(item, list):
            flat_list.extend(flatten_list(item))
        #base case
        else:
            flat_list.append(item)
    return flat_list

def get_earliest_date(df):
    return min(list(map(lambda x : pd.to_datetime(x, format='%d %b %y'),
        set(flatten_list(list(map(lambda x: x.split(", "), df["Date of Hike"]
                                  .unique()))))))).strftime("%d %b %y")

def mmConvertpoint(mm):
    # Convert millimeters to points
    return mm * (72 / 25.4)

def generate_name_tags(name_series, group, reserved=False):
    if group.lower() == "youth":
      group = "Y"
    elif group.lower() == "mentor":
      group = "M"
    else:
      print("Group Can Only Be Either: \"youth\" or \"mentor\" ")
      return "ERROR: INVALID GROUP"

    if not os.path.exists("name_tags_ToBePrinted"):
        os.makedirs("name_tags_ToBePrinted")

    page_width, page_height = A4  # Standard A4 size in points (72 points per inch)
    margin_x, margin_y = 30, 30   # Margins from the edges of the paper in points
    tag_width_mm, tag_height_mm = 85, 55  # Size of each name tag in mm
    tag_width, tag_height = mmConvertpoint(tag_width_mm), mmConvertpoint(tag_height_mm)  # Size of each name tag in points

    # Font size range for the name
    min_name_font_size = 10
    max_name_font_size = 30

    # Existing path for "name_tags_ToBePrinted"
    if reserved:
        pdf_name = master_path + f"name_tags_ToBePrinted/Reserved_UID_{group}.pdf"
    else:
        pdf_name = master_path + f"name_tags_ToBePrinted/name_tags_{group}.pdf"
    c = canvas.Canvas(pdf_name, pagesize=A4)

    # New path for "archive"
    archive_path = master_path + "name_tags_ToBePrinted/archive/"
    if not os.path.exists(archive_path):
        os.makedirs(archive_path)

    format = '%Y-%m-%d %H:%M:%S'
    if reserved:
        archive_pdf_name = archive_path + f"{group}/Reserved_UID_{group}_{pd.Timestamp.today().strftime(format)}.pdf"
    else:
        archive_pdf_name = archive_path + f"{group}/name_tags_{group}_{pd.Timestamp.today().strftime(format)}.pdf"

    x, y = margin_x, page_height - margin_y  # Starting position for the first name tag
    boxes_per_page = 0  # Counter for boxes printed on each page
    rows_per_page = 4  # Number of rows per page
    columns_per_page = 2  # Number of columns per page

    for uid, name in name_series.items():
        # Draw guidelines all around the name tag
        c.setStrokeColorRGB(0.7, 0.7, 0.7)  # Gray color for guidelines
        c.setLineWidth(0.1)
        c.line(x, y, x + tag_width, y)  # Top guideline
        c.line(x, y - tag_height, x + tag_width, y - tag_height)  # Bottom guideline
        c.line(x, y, x, y - tag_height)  # Left guideline
        c.line(x + tag_width, y, x + tag_width, y - tag_height)  # Right guideline

        # Find the maximum font size that fits the name in the available space
        name_font_size = max_name_font_size
        while c.stringWidth(name, "Helvetica-Bold", name_font_size) > tag_width - 10:
            name_font_size -= 1
            if name_font_size < min_name_font_size:
                # If the name still doesn't fit at the smallest font size, break the loop
                break

        # Draw the name in the center both by height and width
        c.setFont("Helvetica-Bold", name_font_size)
        text_width, text_height = c.stringWidth(name, "Helvetica-Bold", name_font_size), name_font_size
        x_offset = (tag_width - text_width) / 2
        y_offset = (tag_height - text_height) / 2  # Center by height
        c.drawString(x + x_offset, y - y_offset, name)

        # Calculate the position of the UID at the bottom right with a margin from the borders
        uid_font_size = 10
        uid_margin_x, uid_margin_y = 5, 5
        uid_x = x + tag_width - uid_margin_x - c.stringWidth("UID: " + group + str(uid), "Helvetica", uid_font_size)
        uid_y = y - tag_height + uid_margin_y

        # Draw the UID at the bottom right in a smaller font
        c.setFont("Helvetica", uid_font_size)
        c.drawString(uid_x, uid_y, "UID: " + group + str(uid))

        # Update x and y position for the next name tag
        boxes_per_page += 1
        if boxes_per_page % columns_per_page == 0:
            # Move to the next row after completing a row
            y -= tag_height
            x = margin_x
        else:
            # Move to the next column in the same row
            x += tag_width

        # Check if the maximum number of rows per page is reached, then start a new page
        if boxes_per_page == rows_per_page * columns_per_page:
            c.showPage()  # Start a new page
            x, y = margin_x, page_height - margin_y
            boxes_per_page = 0  # Reset the boxes count for the new page

    c.save()

def normalize_string(s):
    return s.lower().replace(" ", "")

def create_vcard(df, filename):
    # Open file in write mode
    with open(filename, 'w') as file:
        for index, row in df.iterrows():
            # Create a new vCard
            vcard = vobject.vCard()

            # Add name
            school_company = row['School/Company']
            smc_name = f"SMC {row.full_name} {school_company}"
            vcard.add('n')
            vcard.n.value = vobject.vcard.Name(given=smc_name)

            # Add full name
            vcard.add('fn')
            vcard.fn.value = smc_name

            # Add phone
            tel = vcard.add('tel')
            tel.value = str(row['Whatsapp/mobile Number'])
            tel.type_param = 'CELL'

            # Add company
            qualification_details = row["Qualification"]
            title = row["Major/Title"]
            company_details = f"{qualification_details}, {title}"
            org = vcard.add('org')
            org.value = [company_details]

            # Write vCard to file
            file.write(vcard.serialize())

def process_youth_NAMETAG():
    current_youth_qualification = {'Undergraduate': "UG",
                                   'Postgraduate' : "PG",
                                   'Employed' : "E"}
                                   
    path = master_path + "raw_data/"
    file_names = os.listdir(path)
    file_name = [filename for filename in file_names if "youth" in filename.lower()][0]
    file_path = path + file_name


    df = pd.read_excel(file_path)

    earliest_date = get_earliest_date(df)

    #filter for rows with earliest date
    df = df[df["Date of Hike"].str.contains(earliest_date)]
    df["Date of Hike"] = earliest_date
    df["Qualification"] = df["Undergraduate / Postgraduate / Employed"]\
        .map(lambda x: current_youth_qualification[x])

    df = df[['Date of Hike', 'Full name with CAPS SURNAME (for Name Tag)', 'School/Company',
    'Major/Title', 'Qualification',
    'Year in School or Industry', 'Email', 'Whatsapp/mobile Number']]

    df.columns = ['Date of Hike', 'full_name', 'School/Company',
    'Major/Title', 'Qualification',
    'Year in School or Industry', 'Email', 'Whatsapp/mobile Number']

    df.drop_duplicates(subset=["Email", "Whatsapp/mobile Number"],
                    keep="first",
                    inplace=True)
    df.reset_index(drop=True,
                inplace=True)

    # df.to_csv("database/DB_youth.csv", index=False)
    #Update Repository
    #update last referance date for returning members,
    #update new members to DB
    DB = pd.read_csv(master_path + "database/DB_youth.csv", index_col="UID")
    #update archive first
    format = '%Y-%m-%d %H:%M:%S'
    DB.to_csv(master_path + f"database/archive/DB_youth/DB_youth_{pd.Timestamp.today().strftime(format)}.csv",
              index=True)

    new_df = df.copy()
    temp_new_df = new_df.copy()
    # for row_number, row in temp_new_df.iterrows():

    #     row_email = row["Email"]
    #     row_phone = row["Whatsapp/mobile Number"]

    #     #drop row if email/phone already exists
    #     #update last referance date in DB to today()
    #     if row_email in DB["Email"].values:
    #         new_df = new_df[new_df["Email"] != row_email]
    #         DB.loc[DB["Email"] == row_email, "last_referance"] = pd.Timestamp.today()
    #     elif row_phone in DB["Whatsapp/mobile Number"].values:
    #         new_df = new_df[new_df["Whatsapp/mobile Number"] != row_phone]
    #         DB.loc[DB["Whatsapp/mobile Number"] == row_phone, "last_referance"] = pd.Timestamp.today()

    # Extract unique emails and phone numbers from DB for faster access
    existing_emails = set(DB["Email"])
    existing_phones = set(DB["Whatsapp/mobile Number"])

    # Lists to collect duplicate emails and phones
    duplicate_emails = []
    duplicate_phones = []

    #returning members UID to note
    returning_uid = []
    returning_name = []

    # Iterate through temp_new_df to identify duplicates
    for _, row in temp_new_df.iterrows():
      row_email = row["Email"]
      row_phone = row["Whatsapp/mobile Number"]

      if row_email in existing_emails:
        duplicate_emails.append(row_email)
        name = row["full_name"]
        uid = DB[DB["Email"] == row_email].index
        returning_uid.append(uid)
        returning_name.append(name)
      
      if row_phone in existing_phones:
        duplicate_phones.append(row_phone)
        name = row["full_name"]
        uid = DB[DB["Whatsapp/mobile Number"] == row_phone].index
        returning_uid.append(uid)
        returning_name.append(name)


    # display list of returning uid for reference
    returning_df = pd.DataFrame({"UID": returning_uid,
                                 "Full Name": returning_name
                                 })
    returning_df = returning_df.drop_duplicates(subset='UID', keep='first')
    
    print("Table of Returning Member's UID (Nametag will not be reprinted):")
    display(returning_df)


    # Remove rows with duplicate emails or phones from new_df
    new_df = new_df[~new_df["Email"].isin(duplicate_emails)]
    new_df = new_df[~new_df["Whatsapp/mobile Number"].isin(duplicate_phones)]

    # Update last_referance in DB for the duplicates
    DB.loc[DB["Email"].isin(duplicate_emails), "last_referance"] = pd.Timestamp.today()
    DB.loc[DB["Whatsapp/mobile Number"].isin(duplicate_phones), "last_referance"] = pd.Timestamp.today()


    #update DB index name
    #update new_df index name, set index to UID, update last_referance date
    new_df.reset_index(drop=True,
                    inplace=True)

    if np.isnan(DB.index.max()) or DB.index.max() < 100:
        next_UID = 100
    else:
        next_UID = DB.index.max() + 1

    new_df["UID"] = new_df.index + next_UID
    new_df.set_index("UID",
                    drop=True,
                    inplace=True)
    new_df["last_referance"] = pd.Timestamp.today()

    #add new_df to DB and sort by UID index
    DB = pd.concat([DB, new_df],
                axis=0)
    DB.sort_index(inplace=True)

    #drop index of members not achive for > 6 months,
    #redesignate UID to new members first,
    #if there are underflow, add to excess UID .csv tracker file for next update
    DB["last_referance"] = pd.to_datetime(DB["last_referance"])
    seven_mths_ago = pd.Timestamp.today() - pd.DateOffset(months=7)
    DB = DB[DB["last_referance"] > seven_mths_ago]

    #update DB csv
    DB.to_csv(master_path + "database/DB_youth.csv", index=True)

    '''
    Generate name tags to be printed for new members
    PDF will be saved in directory name_tags_ToBePrinted/
    '''
    generate_name_tags(new_df["full_name"],
                       group="youth")

    # '''
    # Generate contacts to be saved
    # '''
    # # contact_log = pd.read_csv(master_path + "database/contact_log.csv").values.ravel()
    # # new_df = new_df[~new_df["Whatsapp/mobile Number"].isin(contact_log)]


    # ### NUS ###
    # # Create vCard file named 'contacts.vcf'
    # time_format = '%Y-%m-%d %H:%M:%S'
    # #archive update
    # create_vcard(new_df[new_df["School/Company"] == "NUS"], master_path + f'contact/archive/contacts_{pd.Timestamp.today().strftime(time_format)}_NUS.vcf')
    # #save as most recent
    # create_vcard(new_df[new_df["School/Company"] == "NUS"], master_path + f'contact/contacts_youth_NUS.vcf')

    # ### SMU ###
    # # Create vCard file named 'contacts.vcf'
    # time_format = '%Y-%m-%d %H:%M:%S'
    # #archive update
    # create_vcard(new_df[new_df["School/Company"] == "SMU"], master_path + f'contact/archive/contacts_{pd.Timestamp.today().strftime(time_format)}_SMU.vcf')
    # #save as most recent
    # create_vcard(new_df[new_df["School/Company"] == "SMU"], master_path + f'contact/contacts_youth_SMU.vcf')

    # ### Others ###
    # # Create vCard file named 'contacts.vcf'
    # time_format = '%Y-%m-%d %H:%M:%S'
    # #archive update
    # create_vcard(new_df[~((new_df["School/Company"] == "NUS") | (new_df["School/Company"] == "SMU"))], master_path + f'contact/archive/contacts_{pd.Timestamp.today().strftime(time_format)}_OTHERS.vcf')
    # #save as most recent
    # create_vcard(new_df[~((new_df["School/Company"] == "NUS") | (new_df["School/Company"] == "SMU"))], master_path + f'contact/contacts_youth_OTHERS.vcf')
    #######################################################
    ###############
    ###########
    # #update contact log
    # contact_log_df = pd.DataFrame(contact_log, columns=["Whatsapp/mobile Number"])
    # combined_df = pd.concat([contact_log_df, new_df[["Whatsapp/mobile Number"]]], ignore_index=True)



    # '''
    # autogenerate welcome email for new members
    # '''
    # sender= "Wilfred"
    # hike_date = "29th July"
    # location = "Punggol MRT, Exit A"
    # time = "9.20am"
    # message_lst = []

    # for name in df_message["full_name"]:
    #     message = f"""
    #                 Hi *{name}*, this is *{sender}*, the student coordinator for *SMC*.
    #                 For the *{hike_date}* SMC Youths Hiking Library event, do be reminded to be at
    #                 *{location + " " + "by" + " " + time}*. We will carry on with the event regardless of rain or shine.
    #                 The finalised participants' list and pairing table will also be sent to you on Sat morning.

    #                 Do kindly "follow" *SMC LinkedIn* at  https://www.linkedin.com/company/smcmentorship/
    #                 for future events' announcements and publication of youths' reflection articles and feedback.
    #                 Please feel free to reach out if you have any queries as well. Thank you!

    #                 *Kindly acknowledge and confirm.*
    #                 """
    # #insert function for this(TODO)

    print("Youth Processed")
    return None


def process_mentor_NAMETAG():

    path = master_path + "raw_data/"
    file_names = os.listdir(path)
    file_name = [filename for filename in file_names if "mentor" in filename.lower()][0]
    file_path = path + file_name


    df = pd.read_excel(file_path)
    earliest_date = get_earliest_date(df)

    #filter for rows with earliest date
    df = df[df["Date of Hike"].str.contains(earliest_date)]
    df["Date of Hike"] = earliest_date

    df = df[['Date of Hike', 'Full name with CAPS SURNAME (for Name Tag)', 'Company',
    'Title','Industry', 'Email', 'Whatsapp/mobile Number']]
    df.columns = ['Date of Hike', 'full_name', 'Company',
    'Title','Industry', 'Email', 'Whatsapp/mobile Number']

    df.drop_duplicates(subset=["Email", "Whatsapp/mobile Number"],
                    keep="first",
                    inplace=True)
    df.reset_index(drop=True,
                inplace=True)

    # df.to_csv("database/DB_mentor.csv", index=False)
    #Update Repository
    #update last referance date for returning members,
    #update new members to DB
    DB = pd.read_csv(master_path + "database/DB_mentor.csv", index_col="UID")
    #update archive first
    format = '%Y-%m-%d %H:%M:%S'
    DB.to_csv(master_path + f"database/archive/DB_mentor/DB_mentor_{pd.Timestamp.today().strftime(format)}.csv",
              index=True)

    new_df = df.copy()
    temp_new_df = new_df.copy()
    # temp_new_df = new_df.copy()
    # for row_number, row in temp_new_df.iterrows():

    #     row_email = row["Email"]
    #     row_phone = row["Whatsapp/mobile Number"]

    #     #drop row if email/phone already exists
    #     #update last referance date in DB to today()
    #     if row_email in DB["Email"].values:
    #         new_df = new_df[new_df["Email"] != row_email]
    #         DB.loc[DB["Email"] == row_email, "last_referance"] = pd.Timestamp.today()
    #     elif row_phone in DB["Whatsapp/mobile Number"].values:
    #         new_df = new_df[new_df["Whatsapp/mobile Number"] != row_phone]
    #         DB.loc[DB["Whatsapp/mobile Number"] == row_phone, "last_referance"] = pd.Timestamp.today()

    # Extract unique emails and phone numbers from DB for faster access
    existing_emails = set(DB["Email"])
    existing_phones = set(DB["Whatsapp/mobile Number"])

    # Lists to collect duplicate emails and phones
    duplicate_emails = []
    duplicate_phones = []

    #returning members UID to note
    returning_uid = []
    returning_name = []

    # Iterate through temp_new_df to identify duplicates
    for _, row in temp_new_df.iterrows():
      row_email = row["Email"]
      row_phone = row["Whatsapp/mobile Number"]

      if row_email in existing_emails:
        duplicate_emails.append(row_email)
        name = row["full_name"]
        uid = DB[DB["Email"] == row_email].index
        returning_uid.append(uid)
        returning_name.append(name)
      
      if row_phone in existing_phones:
        duplicate_phones.append(row_phone)
        name = row["full_name"]
        uid = DB[DB["Whatsapp/mobile Number"] == row_phone].index
        returning_uid.append(uid)
        returning_name.append(name)


    # display list of returning uid for reference
    returning_df = pd.DataFrame({"UID": returning_uid,
                                 "Full Name": returning_name
                                 })
    returning_df = returning_df.drop_duplicates(subset='UID', keep='first')
    
    print("Table of Returning Member's UID (Nametag will not be reprinted):")
    display(returning_df)


    # Remove rows with duplicate emails or phones from new_df
    new_df = new_df[~new_df["Email"].isin(duplicate_emails)]
    new_df = new_df[~new_df["Whatsapp/mobile Number"].isin(duplicate_phones)]

    # Update last_referance in DB for the duplicates
    DB.loc[DB["Email"].isin(duplicate_emails), "last_referance"] = pd.Timestamp.today()
    DB.loc[DB["Whatsapp/mobile Number"].isin(duplicate_phones), "last_referance"] = pd.Timestamp.today()


    #update DB index name
    #update new_df index name, set index to UID, update last_referance date
    new_df.reset_index(drop=True,
                    inplace=True)

    if np.isnan(DB.index.max()) or DB.index.max() < 100:
        next_UID = 100
    else:
        next_UID = DB.index.max() + 1

    new_df["UID"] = new_df.index + next_UID
    new_df.set_index("UID",
                    drop=True,
                    inplace=True)
    new_df["last_referance"] = pd.Timestamp.today()

    #add new_df to DB and sort by UID index
    DB = pd.concat([DB, new_df],
                axis=0)
    DB.sort_index(inplace=True)

    #drop index of members not achive for > 6 months,
    #redesignate UID to new members first,
    #if there are underflow, add to excess UID .csv tracker file for next update
    DB["last_referance"] = pd.to_datetime(DB["last_referance"])
    seven_mths_ago = pd.Timestamp.today() - pd.DateOffset(months=7)
    DB = DB[DB["last_referance"] > seven_mths_ago]

    #update DB csv
    DB.to_csv(master_path + "database/DB_mentor.csv", index=True)

    '''
    Generate name tags to be printed for new members
    PDF will be saved in directory name_tags_ToBePrinted/
    '''
    generate_name_tags(new_df["full_name"],
                       group="mentor")

    print("Mentor Processed")

    return None

def Reserved_unique_ID_print(group):
    if group.lower() == "youth":
        pass
    elif group.lower() == "mentor":
        pass
    else:
        print("Group Can Only Be Either: \"youth\" or \"mentor\" ")
        return "ERROR: INVALID GROUP"

    if group.lower() == "mentor":
        df = pd.read_csv(master_path + "database/Reserved_UID_mentor.csv", index_col="UID")

    elif group.lower() == "youth":
        df = pd.read_csv(master_path + "database/Reserved_UID_youth.csv", index_col="UID")


    '''
    Generate name tags to be printed for new members
    PDF will be saved in directory name_tags_ToBePrinted/
    '''
    generate_name_tags(df["full_name"],
                       group=group,
                       reserved=True)

    print("Reserved UID Processed")
    return None


def REMOVE_uid_from_DB(uid,
                       group):
  if group.lower() == "youth":
    pass
  elif group.lower() == "mentor":
    pass
  else:
    print("Group Can Only Be Either: \"youth\" or \"mentor\" ")
    return "ERROR: INVALID GROUP"

  DB = pd.read_csv(master_path + f"database/DB_{group}.csv", index_col="UID")
  #update archive first
  format = '%Y-%m-%d %H:%M:%S'
  DB.to_csv(master_path + f"database/archive/DB_{group}/DB_{group}_{pd.Timestamp.today().strftime(format)}.csv",
            index=True)
  
  DB = DB[DB.index != uid]
  #update DB csv
  DB.to_csv(master_path + "database/DB_{group}.csv", index=True)
  return "UID Successfully Removed from DB"


