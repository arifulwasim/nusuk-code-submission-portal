NUSUK CODE SUBMISSION PORTAL

লোকাল Run:
1. Python install করুন।
2. এই folder VS Code-এ Open করুন।
3. Terminal-এ লিখুন: pip install -r requirements.txt
4. তারপর লিখুন: python app.py
5. Browser-এ যান: http://127.0.0.1:5000

Admin:
- Login URL: http://127.0.0.1:5000/admin/login
- Username: admin
- Local default password: ChangeThisAdminPassword123!
- Public deploy করার আগে অবশ্যই ADMIN_PASSWORD environment variable-এ শক্তিশালী password দিন।
- SECRET_KEY-ও অবশ্যই বদলাবেন।

Multiple code submission:
- Text box-এ প্রতিটি code আলাদা লাইনে লিখুন।
- প্রতিটি non-empty line আলাদা serial/count পাবে।
- ফাঁকা line বাদ যাবে।
- Duplicate code-ও আলাদা submission হিসেবে count হবে।

Admin:
- Admin login করে /admin dashboard-এ সব code দেখতে পারবেন।
- Edit দিয়ে code পরিবর্তন করা যাবে।
- Delete দিয়ে code মুছে ফেলা যাবে।
- Edit/Delete-এর পর Excel file-ও update হবে।

GitHub + Render:
1. GitHub repository-তে app.py ও requirements.txt upload করুন।
2. Render-এ New Web Service করে repository connect করুন।
3. Build Command: pip install -r requirements.txt
4. Start Command: gunicorn app:app
5. Environment Variables:
   SECRET_KEY = long-random-secret
   ADMIN_PASSWORD = your-strong-password
   DATA_DIR = /var/data
6. Render Persistent Disk যোগ করুন, Mount Path: /var/data।
   Persistent disk ছাড়া local SQLite/Excel data deploy/restart-এ হারাতে পারে।
7. GitHub Pages Flask backend চালাতে পারে না; GitHub source hosting, Render app hosting করবে।

Security note:
- Admin password কখনো public GitHub code-এ লিখবেন না; environment variable ব্যবহার করুন।
- Public user submission-এর জন্য rate limiting/spam protection ভবিষ্যতে যোগ করা ভালো।
