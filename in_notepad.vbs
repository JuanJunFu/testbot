Set wshshell = wscript.CreateObject("WScript.Shell") 

For i = 1 to 5
    Wshshell.run "Notepad" 
    wscript.sleep 1000 '增加睡眠時間以確保記事本完全打開
    wshshell.sendkeys "S"
    wscript.sleep 100
    wshshell.sendkeys "u"
    wscript.sleep 100
    wshshell.sendkeys " b"
    wscript.sleep 100
    wshshell.sendkeys "s"
    wscript.sleep 100
    wshshell.sendkeys "c"
    wscript.sleep 100
    wshshell.sendkeys "r"
    wscript.sleep 100
    wshshell.sendkeys "i"
    wscript.sleep 100
    wshshell.sendkeys " b"
    wscript.sleep 100
    wshshell.sendkeys "e"
    wscript.sleep 100
    wshshell.sendkeys ""
    wscript.sleep 100
    wshshell.sendkeys " H"
    wscript.sleep 100
    wshshell.sendkeys "I"
    wscript.sleep 100
    wshshell.sendkeys "D"
    wscript.sleep 100
    wshshell.sendkeys "E"
    wscript.sleep 100
    wshshell.sendkeys " "
    wscript.sleep 100
    wshshell.sendkeys "L"
    wscript.sleep 100
    wshshell.sendkeys "A"
    wscript.sleep 100
    wshshell.sendkeys "B"
    wscript.sleep 100
    wshshell.sendkeys ""
    wscript.sleep 100
    wshshell.sendkeys " Y"
    wscript.sleep 100
    wshshell.sendkeys "o"
    wscript.sleep 100
    wshshell.sendkeys "u"
    wscript.sleep 100
    wshshell.sendkeys ""
    wscript.sleep 100
    wshshell.sendkeys "T"
    wscript.sleep 100
    wshshell.sendkeys "u"
    wscript.sleep 100
    wshshell.sendkeys "b"
    wscript.sleep 100
    wshshell.sendkeys "e"
    wscript.sleep 100
    wshshell.sendkeys ""
    wscript.sleep 100
    wshshell.sendkeys "C"
    wscript.sleep 100
    wshshell.sendkeys "h"
    wscript.sleep 100
    wshshell.sendkeys "a"
    wscript.sleep 100
    wshshell.sendkeys "n"
    wscript.sleep 100
    wshshell.sendkeys "n"
    wscript.sleep 100
    wshshell.sendkeys "e"
    wscript.sleep 100
    wshshell.sendkeys "l"
    wscript.sleep 100
Next
