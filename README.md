# 📧 Excel–Outlook Email Automation (VBA)

A fully automated solution to send bulk emails directly from Excel using Outlook.  
Supports personalized messages, attachments, HTML formatting, and automatic status tracking.

---

## 🚀 Features

- Send bulk emails from Excel using Outlook
- Personalized subject & body for each recipient
- Optional file attachments for each row
- Uses Outlook default signature automatically
- HTML email formatting supported
- Real‑time status tracking (Sent / Failed)
- Error‑proof execution with validations
- Clean, modular VBA code

---

## 📂 Project Files

```
Excel-Outlook-Email-Automation-VBA/
│
├── BulkEmailAutomation.bas        # Main VBA module
├── EmailDataTemplate.xlsx         # Excel template for recipient data
├── README.md                      # Project documentation
└── LICENSE                        # MIT License
```

---

## 📝 How It Works

1. Open **EmailDataTemplate.xlsx**
2. Fill in:
   - Recipient Name
   - Recipient Email  
   - Subject  
   - Body  
   - Attachment Path (optional)
3. Import `BulkEmailAutomation.bas` into Excel VBA
4. Run the macro:
   **Developer → Macros → SendBulkEmails**
5. Outlook automatically sends emails  
6. Status updates in Excel (Sent / Pending)

---

## ▶️ Usage Instructions

### **Step 1 — Open Excel Template**
Enter all recipient details in the provided columns.

### **Step 2 — Import VBA Module**
- Open Excel  
- Press **ALT + F11**  
- Go to **File → Import File**  
- Select `BulkEmailAutomation.bas`

### **Step 3 — Run the Automation**
- Go to **Developer → Macros**  
- Select **SendBulkEmails**  
- Click **Run**

---

## 🛠️ Requirements

- Microsoft Excel (2016 or later)
- Microsoft Outlook (configured account)
- Windows OS
- VBA enabled

---

## 📜 License

This project is licensed under the MIT License.

---

## ⭐ Support

If you find this project helpful, please consider giving it a ⭐ on GitHub.