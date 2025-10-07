' Title: Loading an Icon in Inventor 2025
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/loading-an-icon-in-inventor-2025/td-p/12745530#messageview_0
' Category: api
' Scraped: 2025-10-07T13:22:12.657301

stdole.IPictureDisp standardIconIPictureDisp = null;
            stdole.IPictureDisp largeIconIPictureDisp = null;
            if (standardIcon != null)
            {
#pragma warning disable CS0618 // Type or member is obsolete
                standardIconIPictureDisp = Support.IconToIPicture(standardIcon) as stdole.IPictureDisp;
                largeIconIPictureDisp = (stdole.IPictureDisp)Support.IconToIPicture(largeIcon);
#pragma warning restore CS0618 // Type or member is obsolete
            }

            mButtonDef = AddinGlobal.InventorApp.CommandManager.ControlDefinitions.AddButtonDefinition(
                displayName, internalName, commandType,
                clientId, description, tooltip,
                standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);