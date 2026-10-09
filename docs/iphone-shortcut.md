# iPhone Shortcut: "Add to Spooky Season"

Share any Instagram post, Facebook event, web page or flier screenshot straight to your list. About 15 seconds later you get a notification like *"Added: Vampire Disco, Wed Oct 28 🎃"*.

## You'll need

From your sheet: **🎃 Events → iPhone Shortcut details**. It shows your **page link** and your **Shortcut key**.

Keep the key private. Anyone who has it can add events to your list. If it leaks, delete the `SUBMIT_TOKEN` property in **Extensions → Apps Script → Project Settings → Script Properties**, and open **iPhone Shortcut details** again to get a new one.

## Build it (about 5 minutes)

Open the **Shortcuts** app, tap **+**, and name it "Add to Spooky Season".

1. Tap **ⓘ** at the bottom and turn on **Show in Share Sheet**. Set it to receive **Images, URLs and Text**.
2. Add **Get Images from Input** (Shortcut Input).
3. Add **If**: *Images* **has any value**. Inside it, add:
   - **Convert Image** to **JPEG**
   - **Resize Image** to width **1600**
   - **Base64 Encode**, with Line Breaks set to **None**
   - **Set Variable** `Image64`
4. After **End If**, add **Get URLs from Input** (Shortcut Input), then **Set Variable** `Link`.
5. Add **Get Contents of URL**:
   - URL: your **page link**
   - Method: **POST**
   - Request Body: **JSON**, with these fields:
     - `token`: your **Shortcut key**
     - `url`: *Link*
     - `imageBase64`: *Image64*
6. Add **Get Dictionary Value**: `message` from *Contents of URL*.
7. Add **Show Notification** with *Dictionary Value*.

## Use it

In Instagram, Facebook, Safari or Photos: **Share → Add to Spooky Season**.

Screenshots give the best results. A shared Instagram link usually gives a square-cropped flier and no caption, because the post is read without your login.
