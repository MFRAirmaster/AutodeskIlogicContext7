' Title: How to Stop Dialog Box From Popping Up
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-stop-dialog-box-from-popping-up/td-p/13834921
' Category: advanced
' Scraped: 2025-10-09T08:51:25.620013

ThisApplication.SilentOperation = True
ThisApplication.UserInterfaceManager.UserInteractionDisabled = True