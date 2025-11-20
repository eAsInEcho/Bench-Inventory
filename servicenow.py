import time 
import os
import logging
import sys 
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, Toplevel, Menu 
import json

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_properly_sized_dialog(title, min_width=600, min_height=400, parent=None):
    """Create a dialog that's properly sized, positioned, resizable."""
    dialog = tk.Toplevel(parent) if parent else tk.Toplevel()
    dialog.title(title)
    dialog.resizable(True, True) # Make resizable
    dialog.minsize(min_width, min_height) # Set minimum size

    dialog.update_idletasks()
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width - min_width) // 2
    y = (screen_height - min_height) // 2
    dialog.geometry(f"{min_width}x{min_height}+{x}+{y}")

    dialog.lift()
    dialog.focus_force()

    # Configure main grid layout for content + bottom buttons
    dialog.rowconfigure(0, weight=1)
    dialog.rowconfigure(1, weight=0)
    dialog.columnconfigure(0, weight=1)
    return dialog

def add_mousewheel_scrolling(canvas, frame):
    """Add mouse wheel scrolling to a canvas containing a frame."""
    # Ensure canvas is bindable and frame is the content frame
    def _on_mousewheel(event):
        # platform-specific scroll adjustments may be needed
        scroll_units = 0
        if sys.platform.startswith('win'):
            scroll_units = int(-1*(event.delta/120))
        elif sys.platform == 'darwin': # macOS
             scroll_units = int(-1 * event.delta)
        else: # Linux/other
             if event.num == 4: scroll_units = -1
             elif event.num == 5: scroll_units = 1
        if scroll_units != 0:
            canvas.yview_scroll(scroll_units, "units")

    # Bind to the canvas itself, or the frame within it
    # Binding to the frame might feel more natural
    target_widget = frame # Or canvas, depending on desired behavior
    target_widget.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
    target_widget.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
    # For Linux
    target_widget.bind("<Enter>", lambda e: canvas.bind_all("<Button-4>", _on_mousewheel), add='+')
    target_widget.bind("<Enter>", lambda e: canvas.bind_all("<Button-5>", _on_mousewheel), add='+')
    target_widget.bind("<Leave>", lambda e: canvas.unbind_all("<Button-4>"), add='+')
    target_widget.bind("<Leave>", lambda e: canvas.unbind_all("<Button-5>"), add='+')

def create_scrollable_frame(parent):
    """Create a scrollable frame that ensures content is accessible."""
    # Container frame to hold canvas and scrollbar, allows border/padding
    container = ttk.Frame(parent)
    container.pack(fill="both", expand=True) # Make container fill parent

    canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollable_frame = ttk.Frame(canvas) # Frame for the actual content
    canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", tags="scrollable_frame")

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def on_canvas_configure(event):
        # Resize the frame to match canvas width
        canvas.itemconfig(canvas_window, width=event.width)

    scrollable_frame.bind("<Configure>", on_frame_configure)
    canvas.bind("<Configure>", on_canvas_configure)

    # Pack canvas and scrollbar within the container frame
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Add mouse wheel scrolling (ensure this function exists and is imported/defined)
    add_mousewheel_scrolling(canvas, scrollable_frame)

    return scrollable_frame, canvas

def add_context_menu(widget):
    """Add context menu with copy/paste to a widget"""
    menu = Menu(widget, tearoff=0)
    # Simplified check for different widget types
    is_text_or_entry = isinstance(widget, (tk.Text, ttk.Entry, scrolledtext.ScrolledText))

    def popup(event):
        # Update menu state based on selection
        try:
             has_selection = bool(widget.selection_get())
        except tk.TclError:
             has_selection = False
        try:
             can_paste = bool(widget.clipboard_get())
        except tk.TclError:
             can_paste = False

        if is_text_or_entry:
             menu.entryconfig("Cut", state=tk.NORMAL if has_selection else tk.DISABLED)
             menu.entryconfig("Copy", state=tk.NORMAL if has_selection else tk.DISABLED)
             menu.entryconfig("Paste", state=tk.NORMAL if can_paste else tk.DISABLED)
             # Add Select All for Text/Entry widgets
             menu.entryconfig("Select All", state=tk.NORMAL)
        else: # Non-text widgets might not support these actions
             menu.entryconfig("Cut", state=tk.DISABLED)
             menu.entryconfig("Copy", state=tk.DISABLED)
             menu.entryconfig("Paste", state=tk.DISABLED)
             menu.entryconfig("Select All", state=tk.DISABLED)

        try:
            menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            menu.grab_release()

    # Add commands only if widget supports them (basic check)
    if is_text_or_entry:
         menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"), state=tk.DISABLED)
         menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"), state=tk.DISABLED)
         menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"), state=tk.DISABLED)
         menu.add_separator()
         menu.add_command(label="Select All", command=lambda: widget.event_generate("<<SelectAll>>"), state=tk.DISABLED)
         widget.bind("<Button-3>", popup) # Right-click

    return menu

# The JS code for the bookmarklet
js_bookmarklet = """javascript:(function(){try{var mainFrame=document.getElementById('gsft_main')||document.querySelector('iframe[name="gsft_main"]')||document.querySelector('frame[name="gsft_main"]');var frameDoc;if(mainFrame){frameDoc=mainFrame.contentDocument||mainFrame.contentWindow.document;}else{frameDoc=document;}if(!frameDoc){alert('Could not find document. Current page structure: '+document.body.innerHTML.slice(0,500));return;}var data={cmdb_url:window.location.href};var fields=['asset_tag','serial_number','name','host_name','operational_status','install_status','location','u_ci_region','owned_by','assigned_to','manufacturer','model_id','u_model_description','vendor','warranty_expiration','os','os_version','comments'];function findElement(field){var idSelectors=['u_cmdb_ci_notebook.'+field,'sys_original.'+field,'u_cmdb_ci_notebook_'+field,'sys_display.u_cmdb_ci_notebook.'+field];var cssSelectors=['[name="'+field+'"]','[id$="'+field+'"]','[id*="'+field+'"]'];for(var i=0;i<idSelectors.length;i++){var elem=frameDoc.getElementById(idSelectors[i]);if(elem)return elem;}for(var i=0;i<cssSelectors.length;i++){var elems=frameDoc.querySelectorAll(cssSelectors[i]);if(elems.length)return elems[0];}return null;}function getSelectDisplayValue(fieldId){var elem=frameDoc.getElementById(fieldId);if(elem&&elem.tagName==='SELECT'){for(var i=0;i<elem.options.length;i++){if(elem.options[i].value===elem.value){return elem.options[i].text;}}}return null;}function getReferenceDisplayValue(fieldId,fieldName){var displaySpanId='sys_display.'+fieldId;var displaySpan=frameDoc.getElementById(displaySpanId);if(displaySpan){return displaySpan.value||displaySpan.textContent.trim();}var altDisplaySpan=frameDoc.querySelector('input[id$="_'+fieldName+'_display"]');if(altDisplaySpan){return altDisplaySpan.value;}var referenceField=frameDoc.getElementById(fieldId);if(referenceField){var container=referenceField.closest('.container-fluid')||referenceField.closest('td')||referenceField.parentNode;if(container){var displaySpans=container.querySelectorAll('input[type="text"][id*="display"], span[id*="display"]');if(displaySpans.length>0){return displaySpans[0].value||displaySpans[0].textContent.trim();}}}return null;}var data_copy={};for(var i=0;i<fields.length;i++){var fieldId='u_cmdb_ci_notebook.'+fields[i];var elem=frameDoc.getElementById(fieldId);if(elem){var fieldName=fields[i]==='host_name'?%27hostname%27:fields[i]===%27u_model_description%27?%27model_description%27:fields[i]===%27u_ci_region%27?%27ci_region%27:fields[i];if([%27operational_status%27,%27install_status%27].includes(fieldName)){var displayValue=getSelectDisplayValue(fieldId);data_copy[fieldName]=displayValue||elem.value;}else{data_copy[fieldName]=elem.value||elem.textContent||%27%27;}}}for(var i=0;i<fields.length;i++){var fieldId=%27u_cmdb_ci_notebook.%27+fields[i];var fieldName=fields[i]===%27host_name%27?%27hostname%27:fields[i]===%27u_model_description%27?%27model_description%27:fields[i]===%27u_ci_region%27?%27ci_region%27:fields[i];if([%27owned_by%27,%27assigned_to%27,%27location%27,%27manufacturer%27,%27model_id%27,%27vendor%27].includes(fieldName)){var displayValue=getReferenceDisplayValue(fieldId,fields[i]);if(displayValue)data_copy[fieldName]=displayValue;}}delete data_copy.hostname;data=Object.assign(data,data_copy);var resultDiv=document.createElement(%27div%27);resultDiv.style=%27position:fixed;top:10px;left:10px;width:80%;height:80%;z-index:99999;background-color:white;padding:20px;border:2px solid blue;border-radius:5px;box-shadow:0 0 10px rgba(0,0,0,0.5);overflow:auto;font-family:Arial,sans-serif;%27;var header=document.createElement(%27div%27);header.innerHTML=%27<h2>Asset Data Extracted</h2><p>Copy this data to use in the IT Bench Inventory app.</p>%27;resultDiv.appendChild(header);var resultArea=document.createElement(%27textarea%27);resultArea.value=JSON.stringify(data,null,2);resultArea.style=%27width:100%;height:70%;margin:10px 0;font-family:monospace;border:1px solid #ccc;padding:10px;';resultArea.onclick=function(){this.select();};resultDiv.appendChild(resultArea);var buttonContainer=document.createElement('div');buttonContainer.style='margin-top:10px;';var copyButton=document.createElement('button');copyButton.textContent='Copy to Clipboard';copyButton.style='margin-right:10px;padding:8px;background-color:#4CAF50;color:white;border:none;cursor:pointer;';copyButton.onclick=function(){resultArea.select();document.execCommand('copy');alert('Data copied to clipboard!');};buttonContainer.appendChild(copyButton);var closeButton=document.createElement('button');closeButton.textContent='Close';closeButton.style='padding:8px;background-color:#f44336;color:white;border:none;cursor:pointer;';closeButton.onclick=function(){document.body.removeChild(resultDiv);};buttonContainer.appendChild(closeButton);resultDiv.appendChild(buttonContainer);document.body.appendChild(resultDiv);resultArea.select();}catch(e){alert('Error extracting data: '+e.message+'\n\nLine: '+e.lineNumber);}})();"""

def create_bookmark_html(identifier, is_asset=True):
    """Create an HTML file that helps the user create a bookmark"""
    # Generate the URL for reference
    if is_asset:
        url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{identifier}"
    else:
        url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=serial_number%3D{identifier}"
    
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ServiceNow Extractor Bookmark Creator</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.6;
        }}
        h1, h2 {{
            color: #333;
            border-bottom: 1px solid #ddd;
            padding-bottom: 10px;
        }}
        .bookmark {{
            background-color: #f0f0f0;
            border: 1px solid #ccc;
            border-radius: 4px;
            padding: 10px;
            margin: 20px 0;
        }}
        .bookmark-link {{
            display: inline-block;
            background-color: #4CAF50;
            color: white;
            padding: 10px 15px;
            text-decoration: none;
            font-weight: bold;
            border-radius: 4px;
            margin: 10px 0;
        }}
        .instructions {{
            background-color: #fffbd6;
            border-left: 4px solid #ffc107;
            padding: 15px;
            margin: 20px 0;
        }}
        .step {{
            margin-bottom: 20px;
        }}
        img {{
            max-width: 100%;
            border: 1px solid #ddd;
            margin: 10px 0;
        }}
    </style>
</head>
<body>
    <h1>ServiceNow Asset Data Extractor</h1>
    
    <div class="instructions">
        <h2>Create a Bookmark (One-time Setup)</h2>
        <p>Follow these steps to create a bookmark that will extract data from ServiceNow:</p>
        
        <div class="step">
            <h3>Step 1: Create a new bookmark</h3>
            <p>Right-click on your browser's bookmark bar and select "Add new bookmark" or "Add page..." (or press Ctrl+D)</p>
        </div>
        
        <div class="step">
            <h3>Step 2: Enter bookmark details</h3>
            <p>In the bookmark dialog:</p>
            <ul>
                <li>Name: <strong>Extract ServiceNow Data</strong></li>
                <li>URL/Location: <em>Copy the code below</em></li>
            </ul>
        </div>
        
        <div class="bookmark">
            <h3>Copy this entire text as the bookmark URL:</h3>
            <textarea id="bookmarklet-code" rows="3" style="width: 100%; font-family: monospace;">{js_bookmarklet}</textarea>
            <button onclick="copyBookmarklet()">Copy Bookmark Code</button>
        </div>
        
        <div class="step">
            <h3>Step 3: Save the bookmark</h3>
            <p>Click "Save" or "Done" to create the bookmark</p>
        </div>
        
        <div class="step">
            <h3>Step 4: Test the bookmark</h3>
            <p>After saving the bookmark, go to this asset page: <a href="{url}" target="_blank">{url}</a></p>
            <p>Once the ServiceNow page loads, click your new "Extract ServiceNow Data" bookmark</p>
            <p>It should show a popup with the asset data that you can copy</p>
        </div>
    </div>
    
    <div class="instructions">
        <h2>Using the Bookmark</h2>
        <ol>
            <li>Navigate to a ServiceNow asset page</li>
            <li>Click the "Extract ServiceNow Data" bookmark</li>
            <li>Click "Copy to Clipboard" in the popup</li>
            <li>Return to the IT Bench Inventory application and paste the data</li>
        </ol>
    </div>
    
    <script>
        function copyBookmarklet() {{
            var codeTextarea = document.getElementById('bookmarklet-code');
            codeTextarea.select();
            document.execCommand('copy');
            alert('Bookmark code copied! Now create a new bookmark and paste this as the URL.');
        }}
    </script>
</body>
</html>
"""
    
    # Create the HTML file
    filename = "servicenow_bookmark_creator.html"
    with open(filename, "w") as f:
        f.write(html_content)
    
    logger.info(f"Created bookmark HTML file: {filename}")
    
    return filename, url

# --- process_json_data (unchanged logic) ---
def process_json_data(json_data):
    """Process the pasted JSON data and convert it to the expected format"""
    try:
        data = json.loads(json_data)
        required_keys = ['asset_tag', 'serial_number'] # Basic check
        for key in required_keys:
            if key not in data:
                return None, f"Missing required field: {key}"

        # Map ServiceNow fields to database fields carefully
        asset_data = {
            'asset_tag': data.get('asset_tag', ''),
            'serial_number': data.get('serial_number', ''),
            # Use 'name' (often system generated) or 'host_name' if available
            'hostname': data.get('hostname', data.get('name', '')),
            'operational_status': data.get('operational_status', ''),
            'install_status': data.get('install_status', ''),
            'location': data.get('location', ''),
            'ci_region': data.get('ci_region', data.get('u_ci_region','')), # Check both possible names
            'owned_by': data.get('owned_by', ''),
            'assigned_to': data.get('assigned_to', ''),
            'manufacturer': data.get('manufacturer', ''),
            'model_id': data.get('model_id', ''),
             # Check both possible names
            'model_description': data.get('model_description', data.get('u_model_description','')),
            'vendor': data.get('vendor', ''),
            'warranty_expiration': data.get('warranty_expiration', ''),
            'os': data.get('os', ''),
            'os_version': data.get('os_version', ''),
            'comments': data.get('comments', ''),
            'cmdb_url': data.get('cmdb_url', '') # Make sure URL is captured
        }
        # Clean up empty strings to be None if necessary for database
        # for key, value in asset_data.items():
        #     if value == '': asset_data[key] = None
        return asset_data, None
    except json.JSONDecodeError as e:
        return None, f"Invalid JSON format: {str(e)}"
    except Exception as e:
        return None, f"Error processing data: {str(e)}"


# --- Modified scrape_servicenow function ---
def scrape_servicenow(identifier, is_asset=True):
    """Enhanced function to scrape ServiceNow data using the bookmarklet approach"""
    logger.info(f"Starting ServiceNow scrape for {'asset' if is_asset else 'serial'}: {identifier}")

    # Generate direct URL to the CMDB record
    if is_asset:
        url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{identifier}"
    else:
        url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=serial_number%3D{identifier}"

    # --- Dialog Creation ---
    # Use minsize instead of fixed geometry
    dialog = create_properly_sized_dialog("ServiceNow Asset Retrieval", min_width=800, min_height=700)

    # --- Main Content Frame (Expanding Row 0) ---
    main_content = ttk.Frame(dialog)
    main_content.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    main_content.rowconfigure(0, weight=1) # Notebook area expands
    main_content.columnconfigure(0, weight=1)

    # --- Notebook ---
    notebook = ttk.Notebook(main_content)
    notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

    # --- Tabs ---
    main_tab = ttk.Frame(notebook)
    notebook.add(main_tab, text=" Main Workflow ")
    manual_tab = ttk.Frame(notebook)
    notebook.add(manual_tab, text=" Manual Entry ")

    # Configure main_tab grid for expansion (specifically row 2 for step 3 frame)
    main_tab.rowconfigure(2, weight=1)
    main_tab.columnconfigure(0, weight=1)

    # --- Main Tab Setup ---
    # Step 1 Frame
    step1_frame = ttk.LabelFrame(main_tab, text="Step 1: Open asset in ServiceNow")
    step1_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
    ttk.Label(step1_frame, text="• Click 'Open CMDB' to launch the asset page.").pack(anchor="w", padx=10, pady=2)
    ttk.Label(step1_frame, text="• Log in to ServiceNow if necessary.").pack(anchor="w", padx=10, pady=2)
    def open_browser(): webbrowser.open(url, new=2); open_button.config(state="disabled"); bookmark_button.config(state="normal"); status_label_var.set("ServiceNow opened...")
    open_button = ttk.Button(step1_frame, text="Open CMDB", command=open_browser)
    open_button.pack(pady=5)

    # Step 2 Frame
    step2_frame = ttk.LabelFrame(main_tab, text="Step 2: Create and use the scrape bookmark")
    step2_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
    ttk.Label(step2_frame, text="• First time only: Click 'Create Scrape Bookmark'.").pack(anchor="w", padx=10, pady=2)
    ttk.Label(step2_frame, text="• Click the 'Extract ServiceNow Data' bookmark in your browser.").pack(anchor="w", padx=10, pady=2)
    ttk.Label(step2_frame, text="• Click 'Copy to Clipboard' in the popup.").pack(anchor="w", padx=10, pady=2)
    def create_bookmark():
        try:
            filename, asset_url = create_bookmark_html(identifier, is_asset)
            if filename:
                 webbrowser.open(f"file://{os.path.abspath(filename)}", new=2)
                 status_label_var.set("Bookmark creator opened...")
            else:
                 messagebox.showerror("Error", "Failed to create bookmark HTML file.", parent=dialog)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create bookmark HTML: {str(e)}", parent=dialog)
    bookmark_button = ttk.Button(step2_frame, text="Create Scrape Bookmark", command=create_bookmark, state="disabled")
    bookmark_button.pack(pady=5)

    # Step 3 Frame (allow expansion)
    step3_frame = ttk.LabelFrame(main_tab, text="Step 3: Paste and process the data")
    step3_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5) # Expands
    step3_frame.rowconfigure(1, weight=1) # Text area row expands
    step3_frame.columnconfigure(0, weight=1) # Content column expands

    ttk.Label(step3_frame, text="• Paste the copied JSON data below:").grid(row=0, column=0, sticky="w", padx=10, pady=(5,2))

    # ScrolledText (allow expansion)
    json_text = scrolledtext.ScrolledText(step3_frame, height=10, wrap="word") # Wrap text
    json_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=5) # Expands
    add_context_menu(json_text)

    # Paste/Clear Buttons Frame
    paste_frame = ttk.Frame(step3_frame)
    paste_frame.grid(row=2, column=0, sticky="w", padx=10, pady=5) # Align left
    def paste_from_clipboard():
        try:
            # Attempt to get clipboard content and check if it looks like JSON
            clip_content = dialog.clipboard_get()
            if clip_content.strip().startswith('{') and clip_content.strip().endswith('}'):
                 json_text.delete("1.0", "end")
                 json_text.insert("1.0", clip_content)
                 status_label_var.set("JSON data pasted...")
            else:
                 messagebox.showwarning("Paste Warning", "Clipboard content doesn't look like JSON data.", parent=dialog)
        except tk.TclError:
             messagebox.showerror("Error", "Could not get data from clipboard.", parent=dialog)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to paste: {str(e)}", parent=dialog)
    ttk.Button(paste_frame, text="Paste from Clipboard", command=paste_from_clipboard).pack(side="left", padx=5)
    ttk.Button(paste_frame, text="Clear", command=lambda: json_text.delete("1.0", "end")).pack(side="left", padx=5)

    # Process Button
    result = [None]
    def process_data():
        json_data = json_text.get("1.0", "end-1c").strip()
        if not json_data: messagebox.showerror("Error", "No data to process.", parent=dialog); return
        asset_data, error = process_json_data(json_data)
        if error: messagebox.showerror("Processing Error", error, parent=dialog); return
        result[0] = asset_data; status_label_var.set("Data processed successfully!"); complete_button.config(state="normal")
    ttk.Button(step3_frame, text="Process Data", command=process_data).grid(row=3, column=0, pady=(5, 10))

    # Status Label
    status_label_var = tk.StringVar(value="Ready. Open CMDB or create bookmark.")
    status_label = ttk.Label(main_tab, textvariable=status_label_var, font=("", 10), wraplength=750) # Wrap long text
    status_label.grid(row=3, column=0, pady=(5, 10), sticky='ew')

    # --- Manual Tab Setup ---
    manual_tab.rowconfigure(1, weight=1) # Scrollable area expands
    manual_tab.columnconfigure(0, weight=1)
    manual_label = ttk.Label(manual_tab, text="Enter asset information manually:")
    manual_label.grid(row=0, column=0, pady=5, sticky='w', padx=10)
    scroll_container = ttk.Frame(manual_tab)
    scroll_container.grid(row=1, column=0, sticky='nsew', padx=5, pady=5)
    scrollable_frame, canvas = create_scrollable_frame(scroll_container) # Create scrollable part

    # Manual Fields (within scrollable_frame)
    fields = [ "asset_tag", "hostname", "serial_number", "operational_status", "install_status",
               "location", "ci_region", "owned_by", "assigned_to", "manufacturer", "model_id",
               "model_description", "vendor", "warranty_expiration", "os", "os_version" ]
    field_labels = {
    "asset_tag": "Asset Tag (Required)",
    "hostname": "Hostname",
    "serial_number": "Serial Number",
    "operational_status": "Operational Status",
    "install_status": "Install Status",
    "location": "Location",
    "ci_region": "CI Region",
    "owned_by": "Owned By",
    "assigned_to": "Assigned To",
    "manufacturer": "Manufacturer",
    "model_id": "Model ID",
    "model_description": "Model Description",
    "vendor": "Vendor",
    "warranty_expiration": "Warranty Expiration",
    "os": "Operating System",
    "os_version": "OS Version"
}
    entries = {}
    for i, field in enumerate(fields):
         ttk.Label(scrollable_frame, text=f"{field_labels.get(field, field.replace('_',' ').title())}:").grid(row=i, column=0, sticky="w", padx=5, pady=3)
         entry = ttk.Entry(scrollable_frame, width=45) # Adjust width as needed
         entry.grid(row=i, column=1, sticky="ew", padx=5, pady=3)
         entries[field] = entry; add_context_menu(entry)
         # Pre-fill identifier
         if field == "asset_tag" and is_asset: entry.insert(0, identifier)
         elif field == "serial_number" and not is_asset: entry.insert(0, identifier)
    scrollable_frame.columnconfigure(1, weight=1) # Allow entry fields to expand width

    ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields), column=0, sticky="nw", padx=5, pady=3)
    comments_text = tk.Text(scrollable_frame, width=45, height=5)
    comments_text.grid(row=len(fields), column=1, sticky="ew", padx=5, pady=3)
    add_context_menu(comments_text)

    # Manual Submit Button (within scrollable_frame)
    def submit_manual():
        asset_data = {field: entries[field].get().strip() for field in fields}
        asset_data['comments'] = comments_text.get("1.0", "end-1c").strip()
        asset_data['cmdb_url'] = url # Add URL automatically
        # Basic validation
        if not asset_data.get('asset_tag') and not asset_data.get('serial_number'):
             messagebox.showerror("Input Required", "Either Asset Tag or Serial Number is required for manual entry.", parent=dialog)
             return
        if not asset_data.get('asset_tag'): # If tag is missing, try to create one? Or enforce it. For now, enforce.
             messagebox.showerror("Input Required", "Asset Tag is required for manual entry.", parent=dialog)
             return
        result[0] = asset_data; dialog.destroy()

    manual_button_frame = ttk.Frame(scrollable_frame)
    manual_button_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=(10, 5))
    ttk.Button(manual_button_frame, text="Submit Manual Data", command=submit_manual).pack()


    # --- Bottom Button Frame (Fixed Row 1) ---
    button_frame = ttk.Frame(dialog)
    button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
    button_frame.columnconfigure(0, weight=1) # Spacer
    def complete(): dialog.destroy()
    complete_button = ttk.Button(button_frame, text="Complete", command=complete, state="disabled")
    complete_button.grid(row=0, column=1, padx=5)
    cancel_button = ttk.Button(button_frame, text="Cancel", command=dialog.destroy)
    cancel_button.grid(row=0, column=2, padx=5)

    # --- Dialog Main Loop ---
    dialog.wait_window()

    # --- Result Handling ---
    if result[0] is None:
        logger.warning("No data processed or submitted, returning None.")
        # Optionally, ask user if they want to cancel or try manual again, but for now return None
        return None # Indicate failure or cancellation

    logger.info(f"Returning asset data: {result[0].get('asset_tag', 'Unknown')}")
    return result[0]


# --- Modified show_manual_entry_form function (Standalone Fallback) ---
# This version is called if scrape_servicenow fails or is bypassed
def show_manual_entry_form(identifier, is_asset, url, parent=None):
    """Show form for manual data entry as a fallback or standalone"""
    logger.info(f"Showing manual entry form (standalone/fallback) for: {identifier}")

    # Use minsize, make resizable
    entry_window = create_properly_sized_dialog("Manual Asset Entry", min_width=700, min_height=700, parent=parent)

    # --- Content Frame (Expanding Row 0) ---
    content_frame = ttk.Frame(entry_window)
    content_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    content_frame.rowconfigure(2, weight=1) # Scrollable area row expands
    content_frame.columnconfigure(0, weight=1)

    # Instructions and View Button
    ttk.Label(content_frame, text="Please enter the asset information manually:", justify="left").grid(row=0, column=0, pady=5, sticky='w', padx=10)
    def open_browser(): webbrowser.open(url)
    ttk.Button(content_frame, text="View in CMDB (Reference)", command=open_browser).grid(row=1, column=0, pady=5, sticky='w', padx=10)

    # --- Scrollable Area Frame (Expanding Row 2) ---
    scroll_container = ttk.Frame(content_frame)
    scroll_container.grid(row=2, column=0, sticky='nsew', pady=5, padx=5)
    scrollable_frame, canvas = create_scrollable_frame(scroll_container)

    # --- Fields (within scrollable_frame) ---
    fields = [ "asset_tag", "hostname", "serial_number", "operational_status", "install_status",
               "location", "ci_region", "owned_by", "assigned_to", "manufacturer", "model_id",
               "model_description", "vendor", "warranty_expiration", "os", "os_version" ]
    field_labels = {
    "asset_tag": "Asset Tag (Required)",
    "hostname": "Hostname",
    "serial_number": "Serial Number",
    "operational_status": "Operational Status",
    "install_status": "Install Status",
    "location": "Location",
    "ci_region": "CI Region",
    "owned_by": "Owned By",
    "assigned_to": "Assigned To",
    "manufacturer": "Manufacturer",
    "model_id": "Model ID",
    "model_description": "Model Description",
    "vendor": "Vendor",
    "warranty_expiration": "Warranty Expiration",
    "os": "Operating System",
    "os_version": "OS Version"
}
    entries = {}
    for i, field in enumerate(fields):
         ttk.Label(scrollable_frame, text=f"{field_labels.get(field, field.replace('_',' ').title())}:").grid(row=i, column=0, sticky="w", padx=5, pady=3)
         entry = ttk.Entry(scrollable_frame, width=45)
         entry.grid(row=i, column=1, sticky="ew", padx=5, pady=3) # sticky='ew'
         entries[field] = entry; add_context_menu(entry)
         # Pre-fill identifier
         if field == "asset_tag" and is_asset: entry.insert(0, identifier)
         elif field == "serial_number" and not is_asset: entry.insert(0, identifier)
    scrollable_frame.columnconfigure(1, weight=1) # Allow entry column to expand

    ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields), column=0, sticky="nw", padx=5, pady=3)
    comments_text = tk.Text(scrollable_frame, width=45, height=5) # Min height
    comments_text.grid(row=len(fields), column=1, sticky="ew", padx=5, pady=3) # sticky='ew'
    add_context_menu(comments_text)

    # --- Result Storage and Submit Logic (within scrollable_frame) ---
    result = [None]
    def submit_data():
        asset_data = {field: entries[field].get().strip() for field in fields}
        asset_data['comments'] = comments_text.get("1.0", "end-1c").strip()
        asset_data['cmdb_url'] = url # Add reference URL
        # Validation
        if not asset_data.get('asset_tag'):
             messagebox.showerror("Input Required", "Asset Tag is required.", parent=entry_window)
             return
        result[0] = asset_data; entry_window.destroy()

    manual_button_frame = ttk.Frame(scrollable_frame)
    manual_button_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=(10, 5))
    ttk.Button(manual_button_frame, text="Submit Manual Data", command=submit_data).pack()

    # --- Bottom Button Frame (Fixed Row 1) ---
    button_frame = ttk.Frame(entry_window)
    button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10))
    button_frame.columnconfigure(0, weight=1) # Spacer pushes buttons right
    ttk.Button(button_frame, text="Submit", command=submit_data).grid(row=0, column=1, padx=5) # Re-bind submit here too
    ttk.Button(button_frame, text="Cancel", command=entry_window.destroy).grid(row=0, column=2, padx=5)

    # --- Dialog Main Loop ---
    entry_window.wait_window()

    logger.info(f"Manual entry form returned: {result[0].get('asset_tag', 'None') if result[0] else 'None'}")
    return result[0]
    """Show form for manual data entry"""
    logger.info("Showing manual entry form")
    
    entry_window = create_properly_sized_dialog("Manual Asset Entry", 700, 700)
    
    # Create content frame (will expand)
    content_frame = ttk.Frame(entry_window)
    content_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    
    # Instructions
    ttk.Label(content_frame, text="Please enter the asset information manually:").pack(pady=10)
    
    # Open browser button to reference the asset
    def open_browser():
        webbrowser.open(url)
    
    ttk.Button(content_frame, text="View in CMDB", command=open_browser).pack(pady=10)
    
    # Create scrollable frame
    main_frame = ttk.Frame(content_frame)
    main_frame.pack(fill="both", expand=True)
    
    scrollable_frame, canvas = create_scrollable_frame(main_frame)
    
    # Field entry widgets
    fields = [
        "asset_tag", "hostname", "serial_number", "operational_status",
        "install_status", "location", "ci_region", "owned_by", 
        "assigned_to", "manufacturer", "model_id", "model_description",
        "vendor", "warranty_expiration", "os", "os_version"
    ]
    
    field_labels = {
        "asset_tag": "Asset Tag (Required)",
        "hostname": "Hostname",
        "serial_number": "Serial Number",
        "operational_status": "Operational Status",
        "install_status": "Install Status",
        "location": "Location",
        "ci_region": "CI Region",
        "owned_by": "Owned By",
        "assigned_to": "Assigned To",
        "manufacturer": "Manufacturer",
        "model_id": "Model ID",
        "model_description": "Model Description",
        "vendor": "Vendor",
        "warranty_expiration": "Warranty Expiration",
        "os": "Operating System",
        "os_version": "OS Version"
    }
    
    entries = {}
    
    for i, field in enumerate(fields):
        ttk.Label(scrollable_frame, text=f"{field_labels[field]}:").grid(row=i, column=0, sticky="w", padx=10, pady=5)
        entry = ttk.Entry(scrollable_frame, width=40)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries[field] = entry
        
        # Add context menu
        add_context_menu(entry)
        
        # Pre-fill the identifier
        if field == "asset_tag" and is_asset:
            entry.insert(0, identifier)
        elif field == "serial_number" and not is_asset:
            entry.insert(0, identifier)
    
    # Comments field (multiline)
    ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields), column=0, sticky="nw", padx=10, pady=5)
    comments_text = tk.Text(scrollable_frame, width=40, height=5)
    comments_text.grid(row=len(fields), column=1, padx=10, pady=5)
    
    # Add context menu
    add_context_menu(comments_text)
    
    # Result storage
    result = [None]
    
    # Submit function
    def submit_data():
        asset_data = {field: entries[field].get().strip() for field in fields}
        asset_data['comments'] = comments_text.get("1.0", "end-1c")
        asset_data['cmdb_url'] = url
        
        if not asset_data['asset_tag']:
            messagebox.showerror("Error", "Asset Tag is required")
            return
        
        result[0] = asset_data
        entry_window.destroy()
    
    # Button frame at fixed position at bottom of window
    button_frame = ttk.Frame(entry_window)
    button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
    
    ttk.Button(button_frame, text="Submit", command=submit_data).pack(side="right", padx=5)
    ttk.Button(button_frame, text="Cancel", command=entry_window.destroy).pack(side="right", padx=5)
    
    # Wait for window to close
    entry_window.wait_window()
    
    logger.info(f"Manual entry form returned: {result[0].get('asset_tag', 'None') if result[0] else 'None'}")
    return result[0]