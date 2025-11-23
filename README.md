# PowerPoint-Format Tool kit

A collection of PowerPoint VBA tools for fast slide formatting:

- Multi-slot **shape format painter**: copy 3 formats at most + Paste to single image
- **TextBox format painter** (font, highlight, fill, border, align)
- **TextBox pasting**: paste textbook to each slide / every N slides
- **Link helpers**: quickly create internal slide links
- **Crop helpers**: crop-duplicate helpers 
- **Delete placeholders**: delete placeholder title textbox to pages
- **Shape Gridding helpers**: align the selected shapes to keep constant x and y separations between shapes.

---

## Folder structure

- `src/vba/` — exported VBA modules (`*.bas`) you can import to your own PPTM/PPAM
- `examples/` — demo PPT files before/after using the toolkit
- `docs/images/` — screenshots used in this README

---

## Installation

1. Download `dist/macros.ppam`.
2. In PowerPoint: **File → Options → Add-ins**  
3. Manage: `PowerPoint Add-ins` → `Go...`  
4. Click **Add New...**, choose `macros.ppam`.
5. Restart PowerPoint (if needed). A new tab **macros** should appear.

---

## Main features

### Shape format painter (CopyA/B/C + Paste)

- Select a picture or shape. If there is only one picture in the slide page, selection is not required. 
- `CopyShapeFormatA(B/C)`: Click **CopyShapeFormatA** (B/C) to store its format:
   - position / size
   - rotation / flip
   - crop information
   - fill color & transparency
   - line style & weight
- `PasteShapeFormat`: Select one or multiple shapes and click **PasteShapeFormat**. If there is only one picture in the slide page, selection is not required. 
   - 1 stored format → 1 output shape
   - 2–3 stored formats → the target shape is duplicated and each duplicate gets one format
- `PasteII`: click **PasteII** to paste the stored formats to the first picture (by creation order) on all the slides past the selected page.
   - no need to click an image
- `ClearInfo`: click **ClearInfo** to delete the stored formats.
- Beware of actions that are path dependent in microsoft powerpoint. The default order of action has rotation before filpping in this macro.
- If there are none-picutre-type objects selected, they will be omissed and no errors will be prompted. 

### TextBox format painter (CopyTBFormat / PasteTBFomrat)
- `CopyTBFormat`: select a text, textbox or place the cursor amid the text and click **CopyTBFormat** to copy font name, size, bold/italic, underline, highlight color, alignment, fill & border of the textbox.
- `PasteTBFomrat`: select a text, textbox or place the cursor amid the text and click **CopyTBFormat** to apply the format to one or multiple textboxes.
- If a text is selected or the cursor is placed in between texts, the whole textbox will be marked for formatting (copy/paste).
- If one or more object selected is not text, an error message will be prompted and execution is terminated. 

### TextBox pasting
- `PasteObject`: select a shape and click **PasteObject**. starting from current slide, duplicate the textbox into **each slide** on the same position.
- `PasteObjectII`: select a shape and click **PasteObjectII**. starting from current slide, duplicate the textbox into **every N slides** (N is asked via InputBox).
- If a text is selected or the cursor is placed in between texts, the whole textbox will be marked for duplication.

### Link helpers
- `LinkI`: select a text, textbox or picture and click **LinkI**. You will prompted to enter a number as a slide page number for which the link is created upon your selected shape.
- `LinkII`: select a textbox and click **LinkII**. With the same functionality as LinkI, the additional feature is that a return link will be created on your destination page.
   - A textbox must be selected for LinkII. You will be prompted to enter the name for the return link. The textbox will be created using the same format as your selected textbox, and will be placed at the same exact location, too. 

### Crop helpers
- `CropII`: select a picture and click **CropII**. cropping will be applied to a duplicated image right on top of the original selction.
   - Ideal for multi-cropping to the same image. 
- `AntiCrop`: select a picture and click **AntiCrop**. The picture will be restored to its original size and content.
   - In the current version, rotation, flip, fill, and line info are restored to default, because anticrop is designed to quickly reveal the content hidden by cropping. The function is not meant for production display. 
- If there is only one picture in the active slide, no selection is required for these functions. 

### Delete placeholders
- `Deleteplaceholdertitle`: click **Deleteplaceholdertitle** to delete the empty title textboxes on each page past the active slide (including the active slide). If the title textbox is not empty, they won't be deleted.

### Shape Gridding helpers
- `FineGridding`: select shapes and click **FindGridding** to automatically detect the closest grid by which the shapes appear to be aligned, and apply such fine tuning to keep constant x and y separations between shapes.
  - The separation is counted by position .left and .top, not by the center of the shape or other options. 
---

## Development

- Edit macros inside PowerPoint (`macros.pptm`).
- Export modules to `src/vba/*.bas` to keep them under version control.
- Build a new `macros.ppam` when you want to release a new version.
