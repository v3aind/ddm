
import pandas as pd
import streamlit as st
import io

def process_files(file1, file2, file5):
    if file5 is None:
        st.error("Please upload file5 before proceeding.")
        return  # Stop the function if file5 is missing
    
    # Load input files
    file1_df = pd.read_excel(file1)
    file2_df = pd.read_excel(file2)
    file5_df = pd.read_excel(file5)  # This now runs only if file5 is uploaded

    # Further processing...
    st.success("Files processed successfully!")

# Streamlit app
st.title('DDM File Generator (Roaming SC with Gift)')

file1 = st.file_uploader("Upload Roaming SC Completion (Excel)", type=["xls", "xlsx"])
file2 = st.file_uploader("Upload Product Spec Roaming (Excel)", type=["xls", "xlsx"])
file5 = st.file_uploader("Upload All MRID (Excel)", type=["xls", "xlsx"])  # Ensure this is correctly uploaded

# Function to get ruleset names based on POID
def get_ruleset_names(file5_df, po_id):
    """
    Extracts ruleset short names for a given PO ID based on predefined patterns.
    """
    patterns = ["00PRE00", "00ACT00", "GFPRE00", "GFACT00", "GF00"]
    
    # Filter file5_df to get relevant rows for the given po_id
    po_rulesets = file5_df[file5_df["POID"] == po_id]
    
    # Initialize dictionary to store found values
    ruleset_mapping = {pattern: None for pattern in patterns}
    remaining_rule = None

    for _, row in po_rulesets.iterrows():
        ruleset_name = row["Ruleset_ShortName"]
        
        for pattern in patterns:
            if pattern in ruleset_name:
                ruleset_mapping[pattern] = ruleset_name
                break
        else:
            # If it doesn't match any known pattern, store it as remaining
            remaining_rule = ruleset_name

    # Ensure that all required values are present, and fill missing ones
    ruleset_list = [
        ruleset_mapping.get("00PRE00", ""),
        ruleset_mapping.get("00ACT00", ""),
        remaining_rule or "",  # This is for the 3rd position
        ruleset_mapping.get("GFPRE00", ""),
        ruleset_mapping.get("GFACT00", ""),
        ruleset_mapping.get("GF00", "")
    ]
    
    return ruleset_list

# Function to process the uploaded files and provide download link
def process_files(file1, file2, file5):
    if file1 is not None and file2 is not None and file5 is not None:
        # Load input files
        file1_df = pd.read_excel(file1)
        file2_df = pd.read_excel(file2)
        file5_df = pd.read_excel(file5)

        # Validate required columns
        required_columns_file2 = ["Keywords", "Shortcode", "Unreg", "Keyword Alias1", "Keyword Alias2", "Commercial Name", "SIM Action", "SIM Validity", "Package Validity", "Renewal", "PricePre"]
        for col in required_columns_file2:
            if col not in file2_df.columns:
                st.error(f"Missing required column '{col}' in Product Spec Roaming.xlsx")
                return

        output_file_name = None  # Initialize variable for output file name
        
        for index, row in file2_df.iterrows():
            keyword = row["Keywords"]

            # Get PO ID from file1_df based on some criteria (e.g., matching keyword)
            matching_rows = file1_df.loc[file1_df['Keyword'] == keyword, 'POID']

            if not matching_rows.empty:
                po_id_from_file1 = matching_rows.iloc[0]
                output_file_name = f"Prodef DMP-{po_id_from_file1}.xlsx"

                ruleset_names = get_ruleset_names(file5_df, po_id_from_file1)

                # Create a Pandas ExcelWriter
                with io.BytesIO() as output:
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        # Create "PO-Master" sheet
                        po_master_data = {"PO ID": [po_id_from_file1], "Family": ["ROAMINGSINGLECOUNTRY"], "Family Code": ["RSC"]}
                        po_master_df = pd.DataFrame(po_master_data)
                        po_master_df.to_excel(writer, sheet_name="PO-Master", index=False)

                        # Create "Keyword-Master" sheet
                        keyword_master_data = {
                            "Keyword": [row["Keywords"], row["Keywords"], row["Keywords"], "AKTIF_P26", "AKTIF", row["Unreg"]],
                            "Short Code": [str(int(row["Shortcode"])), "124", "929", "122", "122", "122"],
                            "Keyword Type": ["Master", "Master", "Master", "Dormant", "Dormant", "UNREG"]
                        }
                        keyword_master_df = pd.DataFrame(keyword_master_data)
                        keyword_master_df.to_excel(writer, sheet_name="Keyword-Master", index=False)

                        # Create the "Keyword-Alias" sheet
                        keyword_alias_data = {
                            "Keyword": [
                                row["Keywords"],  # 1st row
                                row["Keywords"],  # 2nd row
                            ],
                            "Short Code": [
                                str(int(row["Shortcode"])),  # 1st row from file2 without .0
                                str(int(row["Shortcode"])),  # 2nd row without .0
                            ],
                            "Keyword Aliases": [
                                row["Keyword Alias1"],  # 1st row
                                row["Keyword Alias2"],  # 2nd row
                            ]
                        }
                        keyword_alias_df = pd.DataFrame(keyword_alias_data)
                        keyword_alias_df.to_excel(writer, sheet_name="Keyword-Alias", index=False)

                        # Create the "Ruleset-Header" sheet
                        ruleset_header_data = {
                            "Ruleset ShortName": [
                                ruleset_names[0],  # 00PRE00
                                ruleset_names[1],  # 00ACT00 
                                ruleset_names[1],  # 00ACT00
                                ruleset_names[2],  # remaining_rule
                                ruleset_names[3],  # GFPRE00
                                ruleset_names[4],  # GFACT00
                                ruleset_names[4],  # GFACT00
                                ruleset_names[5],  # GF00
                            ],
                            "Keyword": [row["Keywords"], "AKTIF_P26", "AKTIF", row["Keywords"],row["Keywords"], "AKTIF_P26", "AKTIF", row["Keywords"]],
                            "Keyword Type": ["", "", "", "", "", "", "", ""],
                            "Commercial Name Bahasa": [
                                row["Commercial Name"], 
                                row["Commercial Name"], 
                                row["Commercial Name"],
                                row["Commercial Name"], 
                                row["Commercial Name"],
                                row["Commercial Name"], 
                                row["Commercial Name"],
                                row["Commercial Name"]
                            ],
                            "Commercial Name English": [
                                row["Commercial Name"], 
                                row["Commercial Name"], 
                                row["Commercial Name"],
                                row["Commercial Name"], 
                                row["Commercial Name"], 
                                row["Commercial Name"], 
                                row["Commercial Name"],
                                row["Commercial Name"]
                            ],
                            "Variant Type": ["00", "00", "00", "00", "GF", "GF", "GF", "GF"],
                            "SubVariant Type": ["PRE00", "ACT00", "ACT00", "00000", "PRE00", "ACT00", "ACT00", "00000"],
                            "SimCard Validity": [
                                row["SIM Action"], 
                                row["SIM Action"], 
                                row["SIM Action"],
                                row["SIM Action"], 
                                row["SIM Action"],
                                row["SIM Action"], 
                                row["SIM Action"],
                                row["SIM Action"]
                            ],
                            "LifeTime Validity": [
                                str(int(row["SIM Validity"])) if pd.notna(row["SIM Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["SIM Validity"])) if pd.notna(row["SIM Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else "",
                                str(int(row["Package Validity"])) if pd.notna(row["Package Validity"]) else ""
                            ],
                            "MaxLife Time": ["360", "360", "360", "360", "360", "360", "360", "360"],
                            "UPCC Package Code": [
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else "",
                                file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].iloc[0] if not file1_df.loc[file1_df['Keyword'] == keyword, 'UPCCCode'].empty else ""
                            ],
                            "Claim Command": ["", "", "", "", "", "", "", ""],
                            "Flag Auto": [
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP"
                            ],
                            "Progression Renewal": ["", "", "", "", "", "", "", ""],
                            "Reminder Group Id": ["GROUP18", "GROUP18", "GROUP18", "GROUP18", "GROUP18", "GROUP18", "GROUP18", "GROUP18"],
                            "Amount": [
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                0,
                                0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                0,
                                0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0
                            ],
                            "Reg Subaction": ["1", "1", "1", "1", "1", "1", "1", "1"]
                        }

                        ruleset_header_df = pd.DataFrame(ruleset_header_data)
                        ruleset_header_df.to_excel(writer, sheet_name="Ruleset-Header", index=False)
                        
                        # Ensure MCC is treated as a string and split by commas
                        mcc_raw = str(row['MCC'])  # Convert MCC to string
                        mcc_values = mcc_raw.split(',')  # Split by commas

                        # Add 'm' prefix to each value and strip any surrounding whitespace
                        mcc_prefixed = ','.join([f"m{mcc.strip()}" for mcc in mcc_values])

                        # Split CC values, prefix each with 'c', and join them back with commas
                        cc_raw = str(row['Country Code'])  # Convert CC to string
                        cc_values = str(row['Country Code']).split(',')
                        cc_prefixed = ','.join([f"c{cc.strip()}" for cc in cc_values])

                        # Create DDM-Rule
                        ddm_rule_data ={
                            "Keyword": [row["Keywords"],row["Keywords"], "AKTIF_P26", "AKTIF", row["Keywords"], row["Keywords"],row["Keywords"],row["Keywords"], "AKTIF_P26", "AKTIF", row["Keywords"], row["Keywords"]],
                            "Ruleset ShortName": [
                                ruleset_names[0], ruleset_names[0],  # 00PRE00, 00PRE00
                                ruleset_names[1], ruleset_names[1],  # 00ACT00, 00ACT00 (repeated)
                                ruleset_names[2], ruleset_names[2],  # remaining_rule, remaining_rule
                                ruleset_names[3], ruleset_names[3],  # GFPRE00, GFPRE00
                                ruleset_names[4], ruleset_names[4],  # GFACT00, GFACT00
                                ruleset_names[5], ruleset_names[5]   # GF00, GF00
                            ],
                            "ACTIVE_SUBS": [""] * 12,
                            "OpIndex":[3,4,1,1,1,2,7,8,2,2,5,6],
                            "SALES_AREA": [""] * 12,
                            "ZONE": [""] * 12,
                            "ORIGIN": [
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                "SDP",
                                "SDP",
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA",
                                "SDP",
                                "SDP",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA"
                            ],
                            "RSC_ChildPO": [
                                "PO_ADO_DOR_AKTIF_P26", "PO_ADO_DOR_AKTIF_P26", "", "", "","","PO_ADO_DOR_AKTIF_P26", "PO_ADO_DOR_AKTIF_P26", "", "", "",""
                            ],
                            "RSC_LOCATION": ["DEFAULT", "DEFAULT", "", "", "DEFAULT", "DEFAULT", "DEFAULT", "DEFAULT", "", "", "DEFAULT", "DEFAULT"],
                            "RSC_DEFAULT_SALES_AREA": [""] * 12,
                            "SUBSCRIBER_TYPE": ["PREPAID,POSTPAID"] * 12,
                            "SM_REGION": [""] * 12,
                            "RSC_MAXMPP": [""] * 12,
                            "RSC_RESERVE_BALANCE": [""] * 12,
                            "DA_204": [""] * 12,
                            "UA_165": [""] * 12,
                            "ORDERTYPE": ["REGISTRATION"] * 12,
                            "GIFT": ["FALSE","FALSE","","","FALSE","FALSE", "TRUE","TRUE","","","TRUE","TRUE"],
                            "RSC_CommercialName": [row["Commercial Name"]] * 12,
                            "ROAMING": [
                                "",
                                "",
                                f"IN|{mcc_prefixed},{cc_prefixed},{str(row['MCC_hex']).replace(' ', '').lower()}",
                                f"IN|{mcc_prefixed},{cc_prefixed},{str(row['MCC_hex']).replace(' ', '').lower()}",
                                f"IN|{str(row['MCC_hex']).replace(' ', '').lower()}",
                                f"IN|{str(row['MCC_hex']).replace(' ', '').lower()}",
                                "",
                                "",
                                f"IN|{mcc_prefixed},{cc_prefixed},{str(row['MCC_hex']).replace(' ', '').lower()}",
                                f"IN|{mcc_prefixed},{cc_prefixed},{str(row['MCC_hex']).replace(' ', '').lower()}",
                                f"IN|{str(row['MCC_hex']).replace(' ', '').lower()}",
                                f"IN|{str(row['MCC_hex']).replace(' ', '').lower()}"
                            ],
                            "ROAMINGFLAG": ["EQ|TRUE", "", "", "", "EQ|TRUE", "", "EQ|TRUE", "", "", "", "EQ|TRUE", ""],
                            "RSC_serviceKeyword": ["", "ActivateIntlRoaming", "", "", "", "ActivateIntlRoaming", "", "ActivateIntlRoaming", "", "", "", "ActivateIntlRoaming"],
                            "RSC_serviceName": ["", "ActivateIntlRoaming", "", "", "", "ActivateIntlRoaming", "", "ActivateIntlRoaming", "", "", "", "ActivateIntlRoaming"],
                            "RSC_serviceProvider": ["", "ICARE", "", "", "", "ICARE", "", "ICARE", "", "", "", "ICARE"],
                            "RSC_BYP_CONSENT_CHANNEL" : [
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                "",
                                "",
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                f"{str(row['Channel-SS']).replace(' ', '')},{str(row['Channel-Trad-NonTrad']).replace(' ', '')}",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA",
                                "",
                                "",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA",
                                "UMB,SMS,LTS,V2MYIM3,CHATBOTWA"
                            ],
                            "RSC_RuleSetName": [
                                "GLOBAL_ELIG_ROAMING_PREACT1",
                                "GLOBAL_ELIG_ROAMING_PREACT1",
                                "GLOBAL_ELIG_ROAMING_PREACT2",
                                "GLOBAL_ELIG_ROAMING_PREACT2",
                                "GLOBAL_ELIG_ROAMING_NORMAL",
                                "GLOBAL_ELIG_ROAMING_NORMAL",
                                "GLOBAL_ELIG_ROAMING_PREACT1",
                                "GLOBAL_ELIG_ROAMING_PREACT1",
                                "GLOBAL_ELIG_ROAMING_PREACT2",
                                "GLOBAL_ELIG_ROAMING_PREACT2",
                                "GLOBAL_ELIG_ROAMING_NORMAL",
                                "GLOBAL_ELIG_ROAMING_NORMAL"
                            ],
                            "PREACT_SUBS": [
                                "",
                                "",
                                f"IN|{ruleset_names[0]}",
                                f"IN|{ruleset_names[0]}",
                                "",
                                "",
                                "",
                                "",
                                f"IN|{ruleset_names[3]}",
                                f"IN|{ruleset_names[3]}",
                                "",
                                ""
                            ]
                        }

                        ddm_rule_df = pd.DataFrame(ddm_rule_data)
                        ddm_rule_df.to_excel(writer, sheet_name="DDM-Rule", index=False)

                        # Create Rules-Price
                        rules_price_data ={
                           "Ruleset ShortName": [
                                ruleset_names[0],
                                ruleset_names[0],
                                ruleset_names[1],
                                ruleset_names[1],
                                ruleset_names[2],
                                ruleset_names[2],
                                ruleset_names[3],
                                ruleset_names[4],
                                ruleset_names[4],
                                ruleset_names[5],
                            ],
                            "Variable Name": ["REGISTRATION"] * 3 + ["DORMANT"] + ["REGISTRATION"] * 4 + ["DORMANT"] + ["REGISTRATION"],
                            "Channel":[
                                row["Channel Free"],
                                "DEFAULT",
                                "DEFAULT",
                                ruleset_names[0],
                                row["Channel Free"],
                                "DEFAULT",
                                "DEFAULT",
                                "DEFAULT",
                                ruleset_names[3],
                                "DEFAULT"
                            ],
                            "Price": [
                                0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                0,
                                "",
                                0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                                0,
                                "",
                                int(float(str(row["PricePre"]).replace(",", ""))) if pd.notna(row["PricePre"]) else 0,
                            ],
                            "SID": [
                                "12200001178102", 
                                "12200001178102", 
                                "12200001178102", 
                                "",
                                "12200001178102", 
                                "12200001178102",
                                "12200001178102", 
                                "12200001178102", 
                                "",
                                "12200001178102" 
                            ],
                            "Resultant Shortname": [""] * 3 + [ruleset_names[0]] + [""] * 4 + [ruleset_names[3]] + [""]
                        }

                        rules_price_df = pd.DataFrame(rules_price_data)
                        rules_price_df.to_excel(writer, sheet_name="Rules-Price", index=False)

                        # Create Rules-Renewal
                        rules_renewal_data = {
                            "Ruleset ShortName": ruleset_names,
                            "PO ID": [po_id_from_file1] * 6,
                            "Flag Auto": [
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP",
                                "NO-KEEP" if row["Renewal"] == "No" else "YES-KEEP"
                            ],
                            "Period": [
                                int(row["Dorman"]) if pd.notna(row["Dorman"]) else 0,
                                int(row["Package Validity"]) if pd.notna(row["Package Validity"]) else 0,
                                int(row["Package Validity"]) if pd.notna(row["Package Validity"]) else 0,
                                int(row["Dorman"]) if pd.notna(row["Dorman"]) else 0,
                                int(row["Package Validity"]) if pd.notna(row["Package Validity"]) else 0,
                                int(row["Package Validity"]) if pd.notna(row["Package Validity"]) else 0
                            ],
                            "Period UOM": ["DAY"] * 6,
                            "Flag Charge": ["TRUE"] * 6,
                            "Flag Suspend": ["FALSE"] * 6,
                            "Suspend Period": [""] * 6,
                            "Suspend UOM": [""] * 6,
                            "Flag Option": ["FALSE"] * 6,
                            "Max Cycle": [1] * 6,
                            "Progression Renewal": [""] * 6,
                            "Reminder Group Id": ["GROUP18"] * 6,
                            "Amount": [""] * 6,
                            "Reg Subaction": ["1"] * 6,
                            "Action Failure": ["DEFAULT"] * 6
                        }
                        rules_renewal_df = pd.DataFrame(rules_renewal_data)
                        rules_renewal_df.to_excel(writer, sheet_name="Rules-Renewal", index=False)

                        # Create Case-Type
                        case_type_data = {
                            "RulesetName": ruleset_names,
                            "Case_Type": ["REGISTRATION, UNREG"] *6
                        }

                        case_type_df = pd.DataFrame(case_type_data)
                        case_type_df.to_excel(writer, sheet_name="Case-Type", index=False)

                        # Create the "Offer-DA" sheet
                        headers = ["PO ID", "Offerid", "DA ID", "Benefit Name", "Value", "Zone"]
                        offer_da_data = []  # Initialize as an empty list to store rows

                        def safe_int(value, default=0):
                            """Convert a value to an integer, returning a default value if conversion fails."""
                            try:
                                # Strip whitespace and convert to integer
                                return int(str(value).strip())
                            except (ValueError, TypeError):
                                return default

                        # Check and add data if quota > 0
                        if safe_int(row.get("Quota", 0)) > 0:
                            offer_da_data.append({
                                "PO ID": po_id_from_file1,
                                "Offerid": "",  # Empty string for Offerid
                                "DA ID": "30100",  # Fixed string "30100"
                                "Benefit Name": "DataRoaming",  # Fixed string "DataRoaming"
                                "Value": safe_int(row["Quota"]) * 1073741824,  # quota * 1 GB in bytes
                                "Zone": "NA",  # Empty string for Zone
                            })

                        # Check and add data if Voice > 0
                        if safe_int(row.get("Voice", 0)) > 0:
                            poid_parts = po_id_from_file1.split("_")  # Split POID by "_"
                            if len(poid_parts) >= 5:  # Ensure there are enough parts
                                package_validity = str(row.get("Package Validity", "")).strip()
                                parentpoid = "PO_ADO_CALLBACKHOME_" + poid_parts[4] + "_" + package_validity + "D"
                                offer_da_data.append({
                                    "PO ID": parentpoid,
                                    "Offerid": "",  # Empty string for Offerid
                                    "DA ID": "30194",  # Assuming a different daid for Voice
                                    "Benefit Name": "VoiceRoamingCallBackHome",  # Fixed string "VoiceRoaming"
                                    "Value": safe_int(row["Voice"]) * 60,  # Voice value times 60 in seconds
                                    "Zone": "NA",  # Empty string for Zone
                                })

                        # Create DataFrame
                        if offer_da_data:  # Only create DataFrame if there's data
                            offer_da_df = pd.DataFrame(offer_da_data)
                        else:  # If no data, create an empty DataFrame with the headers
                            offer_da_df = pd.DataFrame(columns=headers)

                        # Write to Excel
                        offer_da_df.to_excel(writer, sheet_name="Offer-DA", index=False)

                        # Create the "Library AddOn_DA" sheet
                        library_addon_headers = ["Ruleset ShortName", "Parentpoid", "Offerid", "daid", "Benefit Name", "Value", "Zone"]
                        library_addon_da_data = []  # Initialize as an empty list to store rows

                        # Repeat Quota data 6 times with ruleset
                        if safe_int(row.get("Quota", 0)) > 0:
                            quota_value = safe_int(row["Quota"]) * 1073741824  # Convert quota to bytes
                            for i in range(len(ruleset_names)):
                                library_addon_da_data.append({
                                    "Ruleset ShortName": ruleset_names[i],  # Append ruleset suffix to POID
                                    "PO ID": po_id_from_file1,
                                    "Quota Name": "DataRoaming",  # Fixed string "DataRoaming"
                                    "DA ID": "30100",  # Fixed string "30100"
                                    "Internal Description Bahasa": "Kuota Roaming",  # Fixed string "DataRoaming"
                                    "External Description Bahasa": "Kuota Roaming",  # Fixed string "DataRoaming"
                                    "Internal Description English": "Roaming Quota",  # Fixed string "DataRoaming"
                                    "External Description English": "Roaming Quota",  # Fixed string "DataRoaming"
                                    "Visibility": "ON",
                                    "Custom": "SHOW",
                                    "Feature": "",
                                    "Initial Value": quota_value,
                                    "Unlimited Benefit Flag": "",
                                    "Scenario": "Rebuy_Upgrade",
                                    "Attribute Name": "DataMainQuota",
                                    "Action": "",
                                })

                        # Repeat Voice data 6 times with ruleset
                        if safe_int(row.get("Voice", 0)) > 0:
                            poid_parts = po_id_from_file1.split("_")  # Split POID by "_"
                            if len(poid_parts) >= 5:  # Ensure there are enough parts
                                package_validity = str(row.get("Package Validity", "")).strip()
                                parentpoid = "PO_ADO_CALLBACKHOME_" + poid_parts[4] + "_" + package_validity + "D"
                                voice_value = safe_int(row["Voice"]) * 60  # Convert voice value to seconds
                                for i in range(len(ruleset_names)):
                                    library_addon_da_data.append({
                                        "Ruleset ShortName": ruleset_names[i],  # Append ruleset suffix to POID
                                        "Po ID": parentpoid,
                                        "Quota Name": "VoiceRoamingCallBackHome",  # Fixed string "DataRoaming"
                                        "DA ID": "30194",  # Fixed string "30100"
                                        "Internal Description Bahasa": "Kuota Nelp ke IM3 dan TRI",  # Fixed string "DataRoaming"
                                        "External Description Bahasa": "Kuota Nelp ke IM3 dan TRI",  # Fixed string "DataRoaming"
                                        "Internal Description English": "Free Call",  # Fixed string "DataRoaming"
                                        "External Description English": "Free Call",  # Fixed string "DataRoaming"
                                        "Visibility": "ON",
                                        "Custom": "VALUEONLY",
                                        "Feature": "",
                                        "Initial Value": voice_value,
                                        "Unlimited Benefit Flag": "",
                                        "Scenario": "Rebuy_Upgrade",
                                        "Attribute Name": "VoiceRoamingCallBackHome",
                                        "Action": "",
                                    })

                        # Create DataFrame
                        if library_addon_da_data:  # Only create DataFrame if there's data
                            library_addon_da_df = pd.DataFrame(library_addon_da_data)
                        else:  # If no data, create an empty DataFrame with the headers
                            library_addon_da_df = pd.DataFrame(columns=library_addon_headers)

                        # Write to Excel
                        library_addon_da_df.to_excel(writer, sheet_name="Library-Addon-DA", index=False)

                        # Create empty Rules-Messages sheet
                        rules_messages_headers = ["PO ID","Ruleset ShortName","Order Status","Order Type","Sender Address","Channel","Message Content Index","Message Content"]
                        rules_messages_data = []  # Initialize as an empty list to store rows

                       # Create DataFrame
                        if rules_messages_data:  # Only create DataFrame if there's data
                            rules_messages_df = pd.DataFrame(rules_messages_data)
                        else:  # If no data, create an empty DataFrame with the headers
                            rules_messages_df = pd.DataFrame(columns=rules_messages_headers)

                        # Write to Excel
                        rules_messages_df.to_excel(writer, sheet_name="Rules-Messages", index=False)

                        # Create StandAlone sheet
                        standalone_data= {
                            "Ruleset ShortName": [
                                ruleset_names[0], ruleset_names[0],  # 00PRE00, 00PRE00
                                ruleset_names[1], ruleset_names[1],  # 00ACT00, 00ACT00 (repeated)
                                ruleset_names[2], ruleset_names[2],  # remaining_rule, remaining_rule
                                ruleset_names[3], ruleset_names[3],  # GFPRE00, GFPRE00
                                ruleset_names[4], ruleset_names[4],  # GFACT00, GFACT00
                                ruleset_names[5], ruleset_names[5]   # GF00, GF00                          
                            ],
                            "PO ID": [po_id_from_file1] * 12,
                            "Scenarios": [
                                "AddonActivation|AddonGiftActivation|AddonGiftRebuy|AddonRebuy",
                                "AddonUnregistration",
                                "AddonActivation|AddonGiftActivation|AddonGiftRebuy|AddonRebuy",
                                "AddonUnregistration",
                                "AddonActivation|AddonGiftActivation|AddonGiftRebuy|AddonRebuy",
                                "AddonUnregistration",
                                "AddonActivation|AddonGiftActivation|AddonGiftRebuy|AddonRebuy",
                                "AddonUnregistration",
                                "AddonActivation|AddonGiftActivation|AddonGiftRebuy|AddonRebuy",
                                "AddonUnregistration",
                                "AddonActivation|AddonGiftActivation|AddonGiftRebuy|AddonRebuy",
                                "AddonUnregistration"                            ],
                            "Type": ["DA"] * 12,
                            "ID": [str(file1_df.loc[file1_df['Keyword'] == keyword, 'DA Standalone'].iloc[0]) if not file1_df.loc[file1_df['Keyword'] == keyword, 'DA Standalone'].empty else ""] * 12,
                            "Value": [str(i) for i in [1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0]],  # Produces ["1", "0", "0", "0", "0", "0"]
                            "UOM": [str(i) for i in [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]],
                            "Validity": [
                                str(row["Dorman"]),
                                "NO_EXPIRY",
                                str(row["Package Validity"]),
                                "NO_EXPIRY",
                                str(row["Package Validity"]),
                                "NO_EXPIRY",
                                str(row["Dorman"]),
                                "NO_EXPIRY",
                                str(row["Package Validity"]),
                                "NO_EXPIRY",
                                str(row["Package Validity"]),
                                "NO_EXPIRY"
                            ],
                            "Provision Payload Value": [""] * 12,
                            "Payload Dependent Attribute": [""] * 12,
                            "ACTION": ["SET"] * 12
                        }

                        standalone_df=pd.DataFrame(standalone_data)
                        standalone_df.to_excel(writer, sheet_name="Standalone", index=False)

                        # Create Rebuy Association sheet - empty need to populate for each country after MR ID fixed
                        rebuy_association_headers= ["Target PO ID","Target Ruleset ShortName","Target MPP","Target Group","Service Type","Rebuy Price","Allow Rebuy","Rebuy Option","Product Family","Source PO ID","Source Ruleset ShortName","Source MPP","Source Group","Vice Versa Consent","Action"]
                        rebuy_association_data = []  # Initialize as an empty list to store rows

                       # Create DataFrame
                        if rebuy_association_data:  # Only create DataFrame if there's data
                            rebuy_association_df = pd.DataFrame(rebuy_association_data)
                        else:  # If no data, create an empty DataFrame with the headers
                            rebuy_association_df = pd.DataFrame(columns=rebuy_association_headers)

                        # Write to Excel
                        rebuy_association_df.to_excel(writer, sheet_name="Rebuy-Association", index=False)

                        # Create UMB Push Category sheet
                        umb_push_category_data= {
                            "Ruleset ShortName": ruleset_names,
                            "Coherence Key": ruleset_names,
                            "Group Category": ["Pkt Internet"] * 6,
                            "Short Code": [str("122")] * 6,
                            "Show Unit": ["SHOW"] * 6,
                            "Action": [""] * 6
                        }

                        umb_push_category_df=pd.DataFrame(umb_push_category_data)
                        umb_push_category_df.to_excel(writer, sheet_name="UMB-Push-Category", index=False)

                    # Move the file pointer to the beginning of the file so it can be downloaded
                    output.seek(0)

                    # Provide a download button for the user
                    st.download_button(
                        label=f"Download {output_file_name}",
                        data=output,
                        file_name=output_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    else:
        st.warning("Please upload all files to proceed.")

# Call the process function if both files are uploaded
if file1 is not None and file2 is not None:
    process_files(file1, file2, file5)
