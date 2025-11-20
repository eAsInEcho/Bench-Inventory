import tkinter as tk
from datetime import datetime

def add_context_menu(widget):
    """Add context menu with copy/paste to a widget"""
    menu = tk.Menu(widget, tearoff=0)
    
    def popup(event):
        try:
            menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            menu.grab_release()
    
    menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"))
    menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
    menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
    
    widget.bind("<Button-3>", popup)  # Right-click
    
    return menu

def add_mousewheel_scrolling(canvas, container_frame):
    """Add mouse wheel scrolling to a canvas with a container frame"""
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _on_frame_enter(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def _on_frame_leave(event):
        canvas.unbind_all("<MouseWheel>")
    
    container_frame.bind("<Enter>", _on_frame_enter)
    container_frame.bind("<Leave>", _on_frame_leave)

def format_timestamp(timestamp_str):
    """Convert ISO timestamp to 'YYYY/MM/DD HH:MM (UTC±H)'."""
    try:
        dt = datetime.fromisoformat(timestamp_str)
 
        # Format date without seconds/microseconds
        formatted_time = dt.strftime("%Y/%m/%d %H:%M")
 
        # Add simplified timezone offset if available
        if dt.tzinfo:
            offset = dt.utcoffset()
            if offset:
                hours = int(offset.total_seconds() // 3600)
                formatted_time += f" (UTC{hours:+})"
 
        return formatted_time
    except Exception:
        return timestamp_str  # fallback if something breaks