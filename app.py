import streamlit as st
import pandas as pd
import io

st.set_page_config(
    layout="wide",
    page_title="Exchange Mailbox Permissions Script Generator",
    page_icon="📧"
)

# Hide only specific Streamlit elements (keep the hamburger menu)
hide_streamlit_style = """
<style>
footer {visibility: hidden;}
.stDeployButton {display: none;}
.css-1dp5vir {display: none;}
[data-testid="stToolbar"] > div:nth-child(2) {display: none;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

def generate_powershell_script(df_selected):
    """Generates the PowerShell script based on the selected DataFrame."""
    powershell_commands = []
    powershell_commands.append("Connect-ExchangeOnline\n")

    for index, row in df_selected.iterrows():
        identity = row['Identity']
        user = row['User']
        action = row['Action']

        if action == 'Add':
            powershell_commands.append(f"Add-MailboxPermission -Identity {identity} -User {user} -AccessRights FullAccess -InheritanceType All -Confirm:$false")
            powershell_commands.append(f"$NameMailBox = Get-EXOMailbox -Identity {identity} | select-object Name")
            powershell_commands.append("$NameMailBox = $NameMailBox -replace '@{Name=',''")
            powershell_commands.append("$NameMailBox = $NameMailBox -replace '}',''")
            powershell_commands.append(f"Add-RecipientPermission -Identity $NameMailBox -Trustee {user} -AccessRights SendAs -Confirm:$false")
        elif action == 'Remove':
            powershell_commands.append(f"Remove-MailboxPermission -Identity {identity} -User {user} -AccessRights FullAccess -InheritanceType All -Confirm:$false")
            powershell_commands.append(f"$NameMailBox = Get-EXOMailbox -Identity {identity} | select-object Name")
            powershell_commands.append("$NameMailBox = $NameMailBox -replace '@{Name=',''")
            powershell_commands.append("$NameMailBox = $NameMailBox -replace '}',''")
            powershell_commands.append(f"Remove-RecipientPermission -Identity $NameMailBox -Trustee {user} -AccessRights SendAs -Confirm:$false")
        powershell_commands.append("") # Add a blank line for readability

    powershell_commands.append("Disconnect-ExchangeOnline -Confirm:$false")
    return "\n".join(powershell_commands)

st.title("Exchange Mailbox Permissions Script Generator")
st.markdown("Upload your CSV file to generate a PowerShell script for managing Exchange mailbox permissions.")

uploaded_file = st.file_uploader("Choose a CSV file", type="csv")

if uploaded_file is not None:
    # Read the CSV file
    try:
        # Use io.BytesIO to read the uploaded file content as a string
        csv_bytes = uploaded_file.getvalue()
        csv_string = csv_bytes.decode('utf-8')
        df = pd.read_csv(io.StringIO(csv_string), delimiter=';')

        if df.empty:
            st.warning("The uploaded CSV file is empty or could not be parsed.")
        else:
            st.subheader("Review and Select Actions")
            st.write("Uncheck any actions you wish to **remove** from the generated script.")

            # Add a 'Select' column for checkboxes
            df['Select'] = True
            
            # Display the DataFrame for selection
            edited_df = st.data_editor(
                df[['Select', 'Identity', 'User', 'Action']],
                column_config={
                    "Select": st.column_config.CheckboxColumn(
                        "Include in Script?",
                        help="Select to include this action in the PowerShell script",
                        default=True,
                    ),
                    "Identity": "Mailbox Identity",
                    "User": "User/Trustee",
                    "Action": "Action Type (Add/Remove)"
                },
                hide_index=True,
                num_rows="dynamic",
            )

            # Filter selected rows
            df_selected_for_script = edited_df[edited_df['Select']]

            if st.button("Generate PowerShell Script"):
                if df_selected_for_script.empty:
                    st.warning("No actions selected to generate the script.")
                else:
                    script_content = generate_powershell_script(df_selected_for_script)
                    st.subheader("Generated PowerShell Script")
                    st.code(script_content, language='powershell')

                    st.download_button(
                        label="Download PowerShell Script",
                        data=script_content,
                        file_name="generated_mailbox_permissions.ps1",
                        mime="application/x-powershell"
                    )
    except Exception as e:
        st.error(f"Error reading CSV file: {e}. Please ensure it's a valid CSV with ';' delimiter.")