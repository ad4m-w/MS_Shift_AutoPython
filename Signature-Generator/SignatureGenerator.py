import tkinter as tk
from tkinter import messagebox
import webbrowser
import os
import subprocess
import shutil

NY_ADDRESS = "One World Trade Center, Suite 49A, New York, New York 10007"
LV_ADDRESS = "3960 Howard Hughes Parkway, Suite 490, Las Vegas, Nevada 89169"
FEEDBACK_LINK = "https://forms.office.com/pages/responsepage.aspx?id=GpUVgFmt_Ei9bXg6_0hYqHB3aPF2iP9LpuHyJKzhq9RUOUU3QjhGNEhaNkJFSDBSV0YyMFNDSEM1VSQlQCN0PWcu&route=shorturl"

def generate_signature_html(first_name, last_name, position="", direct_phone="", ext="", address="", include_meeting_link=True, include_feedback_link=False):
    full_name = f"{first_name} {last_name}".strip()
    image_filename = f"{first_name}{last_name}".replace(" ", "")
    position_html = f"<i>{position}</i>" if position else ""
    direct_phone_html = f"| Direct: {direct_phone}" if direct_phone else ""
    ext_html = f" Ext. {ext} " if ext else ""  # Add extension if provided

    # Feedback link HTML under position if the checkbox is checked
    feedback_html = f'<br><a href="{FEEDBACK_LINK}" style="color: rgb(12,100,192);" target="_blank">Give Us Your Feedback!</a>' if include_feedback_link else ""

    html_content = f"""
<div tabindex="0" dir="ltr" class="K2IBb DziEn RlaWM" role="textbox" aria-multiline="true"
     aria-label="Signature, press Alt+F10 to exit" contenteditable="true"
     style="user-select: text; color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);">

  <p style="text-align: left; margin: 0in;">
    <span style="font-family: Cambria, serif; font-size: 12pt; color: rgb(21, 61, 99);">Best Regards,</span>
  </p>

  <p style="text-align: left; margin: 0in;">
    <span style="font-family: Calibri, sans-serif; font-size: 12pt; color: rgb(23, 78, 134);">&nbsp;</span>
  </p>

  <table cellspacing="0" cellpadding="0" style="background-color: white; width: 812.25pt; color: rgb(36, 36, 36); border-collapse: collapse;">
    <tbody>
      <tr>
        <td style="padding: 0.75pt;">
          <table cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
            <tr>
              <td style="padding: 0.75pt; width: 61.5pt;">
                <img src="https://download.msshift.com/MS_EMAIL_RES/{image_filename}/{image_filename}.png"
                     width="78" height="78" style="width: 0.8229in; height: 0.8229in;">
              </td>
              <td style="padding: 0.75pt; width: 200.75pt;">
                <p style="margin: 0in; font-family: Cambria, serif; font-size: 11pt;">
                  <span style="color: rgb(56, 105, 198);"><b>{full_name}</b></span><br>
                  <span style="color: rgb(21, 61, 99);">{position_html}</span>
                  {feedback_html}  <!-- Feedback link appears here if checked -->
                </p>
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <tr>
        <td style="padding: 0.75pt;">
          <table cellspacing="0" cellpadding="0" style="border-collapse: collapse;">
            <tr>
              <td style="padding: 0.75pt;">
                <p style="margin: 0in; font-family: Cambria, serif; font-size: 11pt;">
                  <span style="color: rgb(56, 105, 198);">MS SHIFT</span><br>
                  <span style="color: black;">{address}</span><br>
                  <span style="color: rgb(117, 117, 117);">Internet:</span>&nbsp;
                  <a href="http://www.msshift.com" style="color: rgb(5, 99, 193);">msshift.com</a><br>
                  <span style="color: rgb(117, 117, 117);">Telephone:</span>
                  <span style="color: rgb(117, 117, 117);">(877) 677-2006 {ext_html}{direct_phone_html}</span>
                  {f'''| <i><a href="https://outlook.office.com/book/MSSHIFTNewClientOnboarding@msshift.com/" style="color: rgb(12,100,192);" target="_blank">Click Here to Meet with Me!</a></i>''' if include_meeting_link else ""}  
                  <br><br>
                </p>
              </td>
            </tr>

            <tr>
              <td style="padding: 0.75pt;">
                <img src="http://msshift.com/images/MS_EMAIL_RES/ms_a_footer_1.gif"
                     width="600" height="138" style="width: 6.25in; height: 1.4375in;"><br> 
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </tbody>
  </table>

  <br>
  <p style="text-align: left; margin: 0in;">
    <span style="font-family: Arial, sans-serif; font-size: 8pt; color: rgb(153, 153, 153);">
      This message is intended only for the use of the addressee and is privileged, confidential
      and protected from disclosure under applicable law. If you receive this message in error
      and are not the intended recipient or the agent responsible for delivering the message to
      the intended recipient, you are hereby notified that any dissemination, distribution or
      copying of this communication is strictly prohibited.
    </span>
  </p>

</div>
"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop, "email_signature.html")

    with open(file_path, "w", encoding="utf-8") as file:
        file.write(html_content)

    return file_path


def open_in_edge(html_file_path):
    edge_path = shutil.which("msedge")
    if not edge_path:
        edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        if not os.path.exists(edge_path):
            edge_path = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"

    if os.path.exists(edge_path):
        subprocess.Popen([edge_path, html_file_path])
    else:
        webbrowser.open(html_file_path)


# GUI
def run_gui():
    def on_submit():
        
        # Make sure there are no empty spaces.
        fname = first_name_entry.get().strip()
        lname = last_name_entry.get().strip()
        pos = position_entry.get().strip()
        phone = phone_entry.get().strip()
        ext = ext_entry.get().strip()  # Get the extension if provided

        # Prevent from working if no first name / last name
        if not fname or not lname:
            messagebox.showwarning("Missing Info", "First and last name are required.")
            return

        if not (ny_var.get() ^ lv_var.get()):
            messagebox.showwarning("Address Required", "Please select one office address.")
            return

        address = NY_ADDRESS if ny_var.get() else LV_ADDRESS
        include_meeting_link = meeting_link_var.get()
        include_feedback_link = feedback_link_var.get()

        html_file = generate_signature_html(fname, lname, pos, phone, ext, address, include_meeting_link, include_feedback_link)
        messagebox.showinfo("Success", "HTML Signature created and opening in Edge.")
        open_in_edge(html_file)

    def toggle_checkbox(source):
        if source == "ny" and ny_var.get():
            lv_var.set(False)
        elif source == "lv" and lv_var.get():
            ny_var.set(False)

    root = tk.Tk()
    root.title("Signature Generator")
    root.geometry("380x400")

    tk.Label(root, text="First Name:").pack(pady=(10, 0))
    first_name_entry = tk.Entry(root)
    first_name_entry.pack()

    tk.Label(root, text="Last Name:").pack()
    last_name_entry = tk.Entry(root)
    last_name_entry.pack()

    tk.Label(root, text="Position (optional):").pack()
    position_entry = tk.Entry(root, width=30)  # Increase width here
    position_entry.pack()

    tk.Label(root, text="Direct Phone (optional):").pack()
    phone_entry = tk.Entry(root)
    phone_entry.pack()

    tk.Label(root, text="Extension (optional):").pack()  # New field for extension
    ext_entry = tk.Entry(root)
    ext_entry.pack()

    tk.Label(root, text="Office Address:").pack(pady=(10, 0))
    ny_var = tk.BooleanVar()
    lv_var = tk.BooleanVar()
    tk.Checkbutton(root, text="New York Office", variable=ny_var, command=lambda: toggle_checkbox("ny")).pack()
    tk.Checkbutton(root, text="Las Vegas Office", variable=lv_var, command=lambda: toggle_checkbox("lv")).pack()

    meeting_link_var = tk.BooleanVar(value=False)  # default: checked
    tk.Checkbutton(root, text='Include "Click Here to Meet with Me!" link', variable=meeting_link_var).pack()

    feedback_link_var = tk.BooleanVar()  # default: unchecked
    tk.Checkbutton(root, text='Include "Give Us Your Feedback!" link', variable=feedback_link_var).pack()

    tk.Button(root, text="Generate Signature", command=on_submit).pack(pady=15)

    root.mainloop()

run_gui()
