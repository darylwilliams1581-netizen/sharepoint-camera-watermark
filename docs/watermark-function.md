# Optional Azure Function — bake watermark into JPEG

Use this when Power Apps captures the photo and you want the stamp in the actual file stored in SharePoint.

POST /api/watermark

```json
{
  "imageBase64": "<base64 jpeg or png>",
  "lines": ["CONFIDENTIAL", "amy@contoso.com", "2026-08-27 08:53:00"]
}
```

See the repo copy under artifacts or implement with Sharp / ImageSharp. Protect the function with a key or Easy Auth and call it only from Power Automate.
