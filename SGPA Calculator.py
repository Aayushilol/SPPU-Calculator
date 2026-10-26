import streamlit as st
import openpyxl
from io import BytesIO

#Created By Aayushi Gavhane :)
#Year_FE(Comp Sci)
# Roll no - 624
# BEE Project


# The title on the streamlit page
st.title("MMIT SGPA Calculater")
st.write("Upload the SPPU Excel file to fill the missing SGPAs for failed students.")

# 1) Taking the file from the prof.
thefile = st.file_uploader("Upload SPPU Excel sheet (.xlsx)", type=["xlsx"])

if thefile:
    # 2) Opening the uploaded File
    wb = openpyxl.load_workbook(thefile)
    sheet = wb.active
    
    st.write("### File Settings")
    # 3) Take the inputs from the user
    starrow = st.number_input("What row does the student data start on?", min_value=1, value=5)
    
    # using .upper() to avoid errors 
    sgpa_col = st.text_input("Enter the SGPA Column Letter (e.g., E):").upper()
    credits_col = st.text_input("Enter the Credits Earned Column Letter (e.g., F):").upper()
    total_credits = st.number_input("Enter Total Semester Credits (e.g., 44):", min_value=1, value=44)
    
    # 4) The execution button
    if st.button("Generate Updated Excel"):
        
        # Checking if the prof. forgot to type the column letters
        if not sgpa_col or not credits_col:
            st.error("Please enter both column letters before generating!")
        else:
            # 5) The main loop processing
            for row in range(starrow, sheet.max_row + 1):
                # Target the exact cell using the prof's typed letter and current row
                sgpa_cell = sheet[f"{sgpa_col}{row}"]
                credits_cell = sheet[f"{credits_col}{row}"]
                
                # Check for the failed student marker,,, need to to check what also comesss!!!!!!
                if sgpa_cell.value == '--':
                    try:
                        # Grab the earned credits, calculate, and round to 2 decimals
                        earned = float(credits_cell.value)
                        newsgpa = (earned / total_credits) * 10
                        sgpa_cell.value = round(newsgpa, 2)
                    except ValueError:
                        # Skip if the credits cell is entirely empty or has weird text means the problem is in credits earned_column
                        pass
            
            # 6) Save the file into temporary invisible RAM (BytesIO)
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success("Successfully calculated and updated all SGPAs!:)")
            
            # 7) Creating the download button with the custom file name
            fris_name = thefile.name
            # Spliting the name and the extension :(
            name_part, extension_part = fris_name.rsplit('.', 1)
            # Combine them with 'ud' 
            # Can change if u don't want ud! just write what you want instead of ud:)
            newnamehehe = f"{name_part}ud.{extension_part}"

            # Create the download button with the dynamic file name;)
            st.download_button(
                label="Download Updated Excel Sheet:)",
                data=output,
                file_name=newnamehehe,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )