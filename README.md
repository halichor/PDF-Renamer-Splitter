# PDF-Renamer-Splitter
A nifty little offline tool to split and rename PDFs based on CSV data, written in Python. [There's a better web-based version, too.](https://github.com/halichor/pdf-split-rename)

## **Downloads**

* The .py files can be found in the src folder of this repo.
* **Get it here for Windows:** [Dropbox](https://www.dropbox.com/scl/fi/t5qc56eoj3aljbajnxioq/Split-Rename-PDF.exe?rlkey=x84ex2cikuoa4s2by1z23os8p&st=7y9v763w&dl=0)

# **Windows (.py Method)**

### **Install Python**

1. Go to this website: [Download Python](https://www.python.org/downloads)  
2. Click the yellow "Download Python" button (it will detect the version for Windows automatically).  
3. **IMPORTANT:** On the installer screen, check the box that says:

   ✅ “Add Python to PATH” before clicking Install Now.

4. Wait for it to install. Done\!

**Notes**

* If you get an error like “Python is not recognized,” you probably forgot to check the “Add to PATH” box when installing. You’ll need to reinstall Python and check that box.

### **Install the Required Tools**

Once Python is installed:

1. Click the **Start menu** (Windows key) and type "**cmd**" — open the **Command Prompt.**  
2. In the black window, copy and paste this command and press Enter:

| `pip install pymupdf pandas tkinterdnd2` |
| :---- |

3. Upgrade your pip version, just in case:

| `python.exe -m pip install --upgrade pip` |
| :---- |

**Notes**

* If pip doesn’t work, try `python -m pip install pymupdf pandas tkinterdnd2`  instead.

### **Run the App**

1. Put your .py file (for example: Split\_Rename\_PDF.py) somewhere like your Desktop.  
2. In the same Command Prompt window, type:

| `cd Desktop` |
| :---- |

3. Then run this:

| `python Split_Rename_PDF.py` |
| :---- |

# **Mac (.app Method)**
*I haven't provided the .app version yet, but you can do it yourself through PyInstall if you'd like.*
1. Unzip the `(Mac) Split_Rename_PDF.zip` file somewhere.  
2. Double-click the file to run it.  
   * If this doesn’t work, right-click the app, then choose **“Open”** (only for the first attempt).  
   * Click **“Open Anyway”** if macOS warns you.  
   * Try the **Unix Executable or .py** methods below if it still doesn't work.

# **Mac (Unix Executable Method)**
*I haven't provided the executable version yet, but you can do it yourself through PyInstall if you'd like.*
1. Just double-click the Split\_Rename\_PDF file (with no file extension) to run it.  
2. A terminal window will pop up first before the app does.  
3. If the above doesn’t work:  
   * Open a **terminal** window and use `cd` to navigate where you saved the file.  
     Ex. `cd ~/Downloads` if it’s in the Downloads folder.  
   * Then run the executable by typing `./Split_Rename_PDF`

# **Mac (.py Method)**

1. Open the **terminal**.  
2. Install Python 3 if you don't have it: [Python Releases for macOS](https://www.python.org/downloads/mac-osx/)  
3. Then install the needed tools:

| `pip3 install pymupdf pandas tkinterdnd2` |
| :---- |

4. Navigate to the file’s folder (assuming it’s in the **Downloads** folder):

| `cd ~/Downloads` |
| :---- |

5. Run the app:

| `python3 Split_Rename_PDF.py` |
| :---- |

# **Troubleshooting**

* **Pip isn’t working? Run the following through Command Prompt or Terminal.**

  `python3 -m ensurepip --default-pip`

  `python3 -m pip install --upgrade pip`

  **For Mac:** `/Library/Frameworks/Python.framework/Versions/3.10/bin/python3.10 -m pip install --upgrade pip`

* **(Mac) Why isn’t the app opening? Why does it take a while to open?**

  Try reopening the app again. And yes, it takes a while. I’m not sure why, but the Unix Executable and .py versions open just fine.

* **(Mac) Why are there two instances of the app?**
  It’s got something to do with PyInstaller. The Unix Executable and .py versions open just one instance in the dock.

* **(Mac) Why is it hard to click/select things?**

  It has something to do with Tkinter (for the GUI) and macOS (Sonoma and/or beyond). You can press **tab** and cycle through the buttons/fields to make them selectable in the meantime, or just be a little patient. 


