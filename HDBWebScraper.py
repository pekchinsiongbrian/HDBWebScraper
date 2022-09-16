import requests
from requests.structures import CaseInsensitiveDict
import xml.etree.ElementTree as et
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog


class HDBWebScraper:

    def __init__(self, master):
        self.master = master
        master.title('HDB Web Scraper')
        master.resizable(False, False)

        self.notif_text = tk.StringVar()
        self.notif_label = tk.Label(master, textvariable=self.notif_text, fg="#FF0000")

        self.file_save_loc_text = tk.StringVar()
        self.file_save_loc_label = tk.Label(master, textvariable=self.file_save_loc_text)

        self.choose_file_save_loc_btn = tk.Button(text="Choose save location", command=self.get_folder_path)
        self.search_btn = tk.Button(text="Search & Generate Sheet", command=self.search)

        self.pcode_label = tk.Label(text="Enter postal code:")
        self.pcode_input = tk.Entry()

        self.notif_label.grid(row=0, column=1, pady=2, padx=50)
        self.file_save_loc_label.grid(row=1, column=1, pady=2, padx=50)
        self.choose_file_save_loc_btn.grid(row=2, column=1, pady=5, padx=50)
        self.pcode_label.grid(row=3, column=1, padx=50)
        self.pcode_input.grid(row=4, column=1, padx=50)
        self.search_btn.grid(row=5, column=1, pady=30, padx=50)

        master.mainloop()

    def get_folder_path(self):
        path = filedialog.askdirectory(title="Select Folder")
        if len(path) > 0:
            self.file_save_loc_text.set(f"{path}")
            self.notif_text.set("")

    def search(self):
        pcode = self.pcode_input.get()
        if pcode.isnumeric() and len(pcode) == 6 and len(self.file_save_loc_text.get()) > 0:
            self.notif_text.set("")

            # Get address
            addr_url = f"https://developers.onemap.sg/commonapi/search?searchVal={pcode}&returnGeom=N&getAddrDetails=N"
            addr_resp = requests.get(addr_url)

            # Set request headers for the next 2 requests
            headers = CaseInsensitiveDict()
            headers["Accept"] = "application/xml"
            headers["Accept-Language"] = "en-US,en;q=0.9"
            headers["Connection"] = "keep-alive"
            headers[
                "Cookie"] = "TS01620995=01e2ec192ae7c095464e0495fde5ba37f3613ba4bb58d08ce955873a4b4abb69a69a25e4a3b2522bfb132aaff6f01d8c8c84b0bc95; _sp_id.1902=426292d1-f844-4fe7-ae16-648697b69c0f.1631196131.1.1631196132.1631196131.a4d24cb0-f61e-4365-9474-b956c28856dd; PD_STATEFUL_9c1319ae-aa5b-11ea-807d-74fe48228f8b=^%^2Fwebapp^%^2FFI10AWCOMMON; PD_STATEFUL_4468c93e-a870-11ea-98b2-74fe48228f8b=^%^2Fwebapp; PD_STATEFUL_c1cfe488-94b2-11e5-8b7f-74fe48068c63=^%^2Fweb; PD_STATEFUL_c1924bfa-94b2-11e5-8b7f-74fe48068c63=^%^2Fweb; TS0173a313=01e2ec192ab8a4c89e7a0599f416ab6a8bda6bb13bb835cdab7b4d84c45e337592df50bd0867e5d1d1c9837c43f7262a2a4313177d; AMWEBJCT^!^%^2FFIM^!JSESSIONID=0000g3uW7jLnd3RM9hNBhe4Dy6N:14be1662-ac54-4d4a-b49a-661408889e57; PD_STATEFUL_682bc6ee-4b76-11e9-8088-74fe48228f8d=^%^2FFIM; PD_STATEFUL_aa070f6c-aa5a-11ea-807d-74fe48228f8b=^%^2Fwebapp^%^2FFI10AWESVCLIST; PD_STATEFUL_06570eb2-9a83-11ea-8e95-74fe48228f8b=^%^2Fweb^%^2Fcommon; PD_STATEFUL_0676188e-9a83-11ea-8e95-74fe48228f8b=^%^2Fweb^%^2Fcommon; dtCookie=v_4_srv_1_sn_613E8E343E1E86C6B5B25E947B0F7010_perc_100000_ol_0_mul_1_app-3A2710b44d13c226a6_1_app-3A7703e52476c3deab_1; PD_STATEFUL_3a219946-97fa-11e5-bed7-74fe48068c63=^%^2FAV02FileUploadServiceWeb; PD_STATEFUL_3a60285a-97fa-11e5-bed7-74fe48068c63=^%^2FAV02FileUploadServiceWeb; PD_STATEFUL_3a40cc44-97fa-11e5-bed7-74fe48068c63=^%^2FAV02FileUploadServiceWeb; TS01170d4d=01e2ec192ab992235f78a7696fa212393b9c9a0af29c895c13f382b626b62e576b6a1da4ea40e4ae1cb7d7413c883c9c52e40fa6f9; AMWEBJCT^!^%^2FFIM^!https^%^3A^%^2F^%^2Fservices2.hdb.gov.sg^%^2FFIM^%^2Fsps^%^2FHDBEserviceProd^%^2Fsaml20FIMSAML20=uuid90712d69-017d-1d45-8706-e3aad63766c4; HDB-RP-S-SESSION-ID=0_lU/8N/S2uzcEFnfu8mXOJ2RGfyTZxOMKALKpTX0isqfLXw5ViCw=; PD_STATEFUL_e307c0e4-bc44-11ea-807e-74fe48228f8b=^%^2Fweb^%^2Fej03; PD_STATEFUL_be3e05d4-bc44-11ea-807e-74fe48228f8b=^%^2Fwebapp^%^2FEJ03WSMapServiceWeb; PD_STATEFUL_be5e9a56-bc44-11ea-807e-74fe48228f8b=^%^2Fwebapp^%^2FEJ03WSMapServiceWeb; PD_STATEFUL_be1eb026-bc44-11ea-807e-74fe48228f8b=^%^2Fwebapp^%^2FEJ03WSMapServiceWeb; JSESSIONIDAIASP1=0000hy74jEVvsj9MkPOQwXe6MXr:1d71i54uh; PD_STATEFUL_00da1fa0-ac72-11ea-807d-74fe48228f8b=^%^2Fwebapp^%^2FBB14LFEESENQ; TS015ddc07=01e2ec192a06bc94b0fcd1fb6f6b6c54a797f257601173ba96f4f1f9d42a37b3dd11eb6f7b6dbaed41db1be85aca0168ab79d9e77a; info-msg-C1B13EF616969A6E482588BE00040630=true; info-msg-2183B92C5216FAED482585110032624D=true; spcptrackingp1=1663204794394__aa6b_3267c05e15e2; JSESSIONIDP1=0000FOgbZMoFk05YhM3w-zd-xQM:19nr4sl2a; HDB-RP-S-SESSION-ID=.3; isg003=^!Jgi120MJtb4cPld+I6OxAHPnOkf+jo2nOuBWgsky9BYQ7obXCk3M7GqnJy9tZMmvGnx7MxNP3Am1Fek=; TS0160c4dd=01e2ec192a2292e6b45bcbb6c65e04c905e6cb9646f66988437f0a9488b9c9c7cdab44a6e289c4c6446dc36c7a3d98afbd0f47067aba708b63eada3922a6fba7ff50f00a88628646ea787eabcd999e81aaf337c3603e917a8cdab20d5b458cdfbbb956fc0c2cd36bd9e8f33898bccfe7caabeb19fa"
            headers["Referer"] = "https://services2.hdb.gov.sg/web/fi10/emap.html"
            headers["Sec-Fetch-Dest"] = "empty"
            headers["Sec-Fetch-Mode"] = "cors"
            headers["Sec-Fetch-Site"] = "same-origin"
            headers[
                "User-Agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
            headers["X-Requested-With"] = "XMLHttpRequest"

            # Get lease details
            lease_url = f"https://services2.hdb.gov.sg/webapp/BB14ALeaseInfo/BB14SGenerateLeaseInfoXML?postalCode={pcode}"
            lease_resp = requests.get(lease_url, headers=headers)

            # Get "resale flat prices of this block" details
            price_details_url = f"https://services2.hdb.gov.sg/webapp/BB33RTISMAP/BB33SResaleTransMap?postal={pcode}"
            price_details_resp = requests.get(price_details_url, headers=headers)

            if addr_resp.status_code != 200 or lease_resp.status_code != 200 or price_details_resp.status_code != 200:
                self.notif_text.set(f"An error occurred (code: {lease_resp.status_code}). Please inform dev.")
            else:
                if len(addr_resp.json()['results']) == 0:
                    self.notif_text.set("Invalid postal code")
                    return
                addr = addr_resp.json()['results'][0]['SEARCHVAL']

                lease_xml_data = lease_resp.text
                lease_xml = et.fromstring(lease_xml_data)
                lease_start_date = ""
                lease_duration = ""
                remaining_lease = ""
                for child in lease_xml:
                    if child.tag == "LeasePeriod":
                        lease_duration = child.text + " years"
                        continue
                    if child.tag == "LeaseCommencedDate":
                        lease_start_date = child.text
                        continue
                    if child.tag == "LeaseRemaining":
                        remaining_lease += child.text + " years "
                        continue
                    if child.tag == "LeaseMonth":
                        remaining_lease += child.text + " months"
                        continue

                price_xml_data = price_details_resp.text
                price_xml = et.fromstring(price_xml_data)
                table_row = []
                price_dict = {}
                for child in price_xml:
                    if child.tag == "Dataset":
                        for grandchild in child:
                            price_dict[grandchild.tag] = grandchild.text
                        data = {"Flat Type": price_dict['flattype'],
                                "Flat Model": price_dict['modldesc'],
                                "Storey": price_dict['numrange'],
                                "Floor Area": price_dict['floorarea'],
                                "Lease Commence Date": price_dict['dteleasecomm'],
                                "Remaining Lease": price_dict['balleasetenure'] + " years " + price_dict[
                                    'balleasetenuremonths'] + " months",
                                "Resale Price": "$" + price_dict['reslprice'],
                                "Resale Registration Date": price_dict['dteregistration']
                                }
                        table_row.append(data)

                df = pd.DataFrame(table_row)

                as_at = datetime.today().strftime('%d-%m-%Y')
                workbook_name = f"{self.file_save_loc_text.get()}/{pcode}.xlsx"
                with pd.ExcelWriter(path=workbook_name, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name=f"{as_at}", startrow=5, index=False)
                    workbook = writer.book
                    worksheet1 = writer.sheets[f"{as_at}"]
                    worksheet1.set_column(0, 7, 20)

                    header_format = workbook.add_format({'bold': True,
                                                         'border': True,
                                                         'align': 'center',
                                                         'valign': 'vcenter',
                                                         'fg_color': '#ADCEFF',
                                                         'font_size': 12,
                                                         'text_wrap': True})
                    total_format = workbook.add_format({'bold': True,
                                                        'border': True,
                                                        'align': 'left',
                                                        'valign': 'vcenter',
                                                        'fg_color': '#FFD08C',
                                                        'font_size': 11})
                    table_title = f"{addr}"
                    worksheet1.merge_range('A1:H1', table_title, header_format)

                    worksheet1.write('A2', "Lease Start Date", total_format)
                    worksheet1.write('A3', "Lease Duration", total_format)
                    worksheet1.write('A4', "Remaining Lease", total_format)
                    worksheet1.write('B2', lease_start_date, total_format)
                    worksheet1.write('B3', lease_duration, total_format)
                    worksheet1.write('B4', remaining_lease, total_format)

                    self.notif_text.set("Sheet generated!")
        else:
            if len(self.file_save_loc_text.get()) == 0:
                self.notif_text.set("Choose a folder to save the generated sheet in!")
            else:
                self.notif_text.set("Invalid postal code!")


# Flat Type | Flat Model | Storey | Floor Area | Lease Commence Date | Remaining Lease | Resale Price | Resale Registration Date
# flattype | modldesc | numrange | floorarea | dteleasecomm | balleasetenure + " years " + balleasetenuremonths + " months" | reslprice | dteregistration


root = tk.Tk()
WebScraper = HDBWebScraper(root)
