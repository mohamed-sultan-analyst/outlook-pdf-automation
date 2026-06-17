import os
import time
import win32com.client
import pyautogui
import pygetwindow as gw
from reportlab.pdfgen import canvas
from PIL import Image
from pypdf import PdfReader, PdfWriter


BASE_PATH = r"E:\New folder\16-DEBIT NOTE"
FINAL_PATH = os.path.join(BASE_PATH, "Processed_Debit_Notes")
TEMP_DIR = os.path.join(BASE_PATH, "Temp_Internal")

for p in [FINAL_PATH, TEMP_DIR]:
    if not os.path.exists(p): 
        os.makedirs(p)

def create_pdf_from_img(img_path, pdf_path):
   
    try:
        img = Image.open(img_path)
        w, h = img.size
        img.close()
        c = canvas.Canvas(pdf_path, pagesize=(w, h))
        c.drawImage(img_path, 0, 0, width=w, height=h)
        c.save()
        return True
    except Exception as e:
        print(f"❌ Error creating PDF from image: {e}")
        return False

def run_automation():
    print("🚀 Logic Update: Closing emails fast & forcing fresh PDF creation...")
    
    
    active_win = gw.getActiveWindow()
    if active_win:
        active_win.minimize()
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6).Folders("Debit Note")
        messages = list(inbox.Items)
        
        for msg in messages:
            try:
                
                if msg.UnRead and "CONFIRM" in (msg.Subject + msg.Body).upper():
                    
                    
                    for f in os.listdir(TEMP_DIR):
                        try: os.remove(os.path.join(TEMP_DIR, f))
                        except: pass

                    
                    msg.Display()
                    time.sleep(2) 
                    
                    email_windows = gw.getWindowsWithTitle(msg.Subject)
                    if email_windows:
                        email_win = email_windows[0]
                        try:
                            email_win.activate()
                        except:
                            
                            pyautogui.press('alt')
                            email_win.activate()
                        
                        email_win.maximize()
                        time.sleep(1)

                    
                    temp_img = os.path.join(TEMP_DIR, "current_email.png")
                    pyautogui.screenshot(temp_img)
                    
                    
                    current_subject = msg.Subject
                    msg.Close(0) 
                    print(f"📧 Captured and Closed: {current_subject}")
                    
                    temp_pdf = os.path.join(TEMP_DIR, "current_email.pdf")
                    if not create_pdf_from_img(temp_img, temp_pdf):
                        continue

                    
                    for att in msg.Attachments:
                        if att.FileName.lower().endswith(".pdf"):
                            dn_number = att.FileName.split('.')[0]
                            local_pdf = os.path.join(BASE_PATH, f"{dn_number}.pdf")
                            
                            if os.path.exists(local_pdf):
                          
                                writer = PdfWriter()
                                
                               
                                writer.add_page(PdfReader(temp_pdf).pages[0])
                                
                                
                                original_reader = PdfReader(local_pdf)
                                for page in original_reader.pages:
                                    writer.add_page(page)
                                
                                output = os.path.join(FINAL_PATH, f"{dn_number}.pdf")
                                
                               
                                if os.path.exists(output):
                                    try: os.remove(output)
                                    except: pass

                                with open(output, "wb") as f:
                                    writer.write(f)
                                
                                print(f"✅ Final PDF Created: {dn_number}.pdf")

                    
                    msg.UnRead = False 
                    
                    del msg

            except Exception as e:
                print(f"⚠️ Error processing email: {e}")
                continue
                
    except Exception as e:
        print(f"❌ Outlook Error: {e}")
    finally:
       
        if active_win:
            active_win.restore()

if __name__ == "__main__":
    run_automation()