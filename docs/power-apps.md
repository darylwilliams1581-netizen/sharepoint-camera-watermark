# Power Apps + SharePoint camera (field photos)

Use Power Apps if staff already open apps from Teams or a SharePoint site.

Limitation: the Power Apps camera control does not write a watermark into the image pixels. Overlay labels are only visible in the app. For a real stamp on the file, send the photo through Power Automate.

## Capture into a collection

```powerfx
Collect(
    colShots,
    {
        Id: GUID(),
        Photo: Camera1.Photo,
        Caption: txtCaption.Text,
        Tag: txtTag.Text,
        TakenOn: Now(),
        UserEmail: User().Email
    }
)
```

## Upload all via flow

```powerfx
Set(varBusy, true);
ForAll(
    colShots,
    WatermarkAndUpload.Run(
        Caption,
        Tag,
        UserEmail,
        Text(TakenOn, "yyyy-mm-dd hh:mm:ss"),
        {
            file: {
                name: "FieldCam.jpg",
                contentBytes: Photo
            }
        }
    )
);
Clear(colShots);
Set(varBusy, false);
Notify("Photos uploaded", NotificationType.Success)
```

Flow trigger: Power Apps (V2) with Caption, Tag, UserEmail, TakenOn, and File. Create the file in a SharePoint library after optional watermarking. Embed the published app with the Power Apps web part on a SharePoint page. Phone users should open Power Apps mobile so the device camera is available.
