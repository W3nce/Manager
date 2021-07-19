import tkinter as tk
from tkinter import ttk

import os
from PIL import Image, ImageTk

from tkinter import messagebox 
import CountryRef

TreeViewSAVED = 0
RFQIssued = 0


IM_CHECKED = os.path.join(os.path.join(os.getcwd(),'icons'), "checked.png")      
IM_UNCHECKED = os.path.join(os.path.join(os.getcwd(),'icons'), "unchecked.png") 




class CbTreeview(ttk.Treeview):
    
    def __init__(self, master=None, **kw):
        
        self.root = master
        kw.setdefault('style', 'cb.Treeview')
        kw.setdefault('show', 'headings')  # hide column #0
        ttk.Treeview.__init__(self, master, **kw)
        # create checheckbox images
        
        self.im_checked = tk.PhotoImage('checked',
                                          file = IM_CHECKED, master=self)
        self.im_unchecked = tk.PhotoImage('unchecked',
                                          file = IM_UNCHECKED, master=self)
# =============================================================================
#         self._im_checked = tk.PhotoImage('checked',
#                                          data=b'GIF89a\x0e\x00\x0e\x00\xf0\x00\x00\x00\x00\x00\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x0e\x00\x0e\x00\x00\x02#\x04\x82\xa9v\xc8\xef\xdc\x83k\x9ap\xe5\xc4\x99S\x96l^\x83qZ\xd7\x8d$\xa8\xae\x99\x15Zl#\xd3\xa9"\x15\x00;',
#                                          master=self)
#         self._im_unchecked = tk.PhotoImage('unchecked',
#                                            data=b'GIF89a\x0e\x00\x0e\x00\xf0\x00\x00\x00\x00\x00\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x0e\x00\x0e\x00\x00\x02\x1e\x04\x82\xa9v\xc1\xdf"|i\xc2j\x19\xce\x06q\xed|\xd2\xe7\x89%yZ^J\x85\x8d\xb2\x00\x05\x00;',
#                                            master=self)
#         
# =============================================================================
        
        style = ttk.Style(self)
        style.configure("cb.Treeview.Heading", font=(None, 9))
        # put image on the right
        style.layout('cb.Treeview.Row',
                     [('Treeitem.row', {'sticky': 'nswe'}),
                      ('Treeitem.image', {'side': 'right', 'sticky': 'e'})])

        # use tags to set the checkbox state
        self.tag_configure('checked', image='checked')
        self.tag_configure('unchecked', image='unchecked')

    def tag_add(self, item, tags):
        new_tags = tuple(self.item(item, 'tags')) + tuple(tags)
        self.item(item, tags=new_tags)

    def tag_remove(self, item, tag):
        tags = list(self.item(item, 'tags'))
        tags.remove(tag)
        self.item(item, tags=tags)

    def insert(self, parent, index, iid=None, **kw):
        item = ttk.Treeview.insert(self, parent, index, iid, **kw)
        self.tag_add(item, (item, 'unchecked'))
        self.tag_bind(item, '<ButtonRelease-1>',
                      lambda event: self._on_click(event, item))
        
        self.tag_bind(item, '<Double-Button-1>',
                      lambda event: self.onDoubleClick(event))
        #self.tag_bind(item, '<Tab>', self.onTabKey)
        
    def _on_click(self, event, item):
        global TreeViewSAVED,RFQIssued

        """Handle click on items."""
        if self.identify_row(event.y) == item:
            if self.identify_column(event.x) == '#16': # click in 'Served' column
                # toggle checkbox image
                if RFQIssued: 
                    if self.tag_has('checked', item):
                        self.tag_remove(item, 'checked')
                        self.tag_add(item, ('unchecked',))
                    else:
                        self.tag_remove(item, 'unchecked')
                        self.tag_add(item, ('checked',))
                    TreeViewSAVED = 0
                else :
                    messagebox.showerror('Confirm Purchase' ,'RFQ not issued!',parent = self.root)            
        
    def onDoubleClick(self, event):
        global TreeViewSAVED,RFQIssued
       
            
        TreeViewSAVED = 0
        ''' Executed, when a row is double-clicked. Opens 
        read-only EntryPopup above the item's column, so it is possible
        to select text '''

        # close previous popups
        # self.destroyPopups()
    
        # what row and column was clicked on
        rowid = self.identify_row(event.y)
        column = self.identify_column(event.x)
        
        FullWidth = self.root.winfo_width()
        EntryWidth = 0.785 * FullWidth - 23
        
        if column == '#13': 
            if RFQIssued:
                pass
            else:
                if messagebox.askyesno('Edit Unit Price' ,'RFQ not issued!\nDo you want to Continue?',parent = self.root):
                    pass
                else:
                    return
        
            # get column position info
            x,y,width,height = self.bbox(rowid, column)
        
            # y-axis offset
            pady = height // 2
    
        
            
        
            # place Entry popup properly         
            text = self.item(rowid, 'values')[12]
            self.entryPopup = EntryPopup(self, rowid, text)
            self.entryPopup.place( x=int(EntryWidth), y=y+pady, anchor=tk.W, width = int(0.102*EntryWidth - 20))
            
    def CheckUncheckAll(self):
        print(RFQIssued)
        if self.get_children():
            pass
        else:
            messagebox.showerror('Confirm Purchase' ,'No Parts in RFQ!',parent = self.root)  
            return
        
        if RFQIssued:
            for line in self.get_children():
                if self.tag_has('checked', line):
                    self.tag_remove(line, 'checked')
                    self.tag_add(line, ('unchecked',))
                else:
                    self.tag_remove(line, 'unchecked')
                    self.tag_add(line, ('checked',))
        else :
            messagebox.showerror('Confirm Purchase' ,'RFQ not issued!',parent = self.root)   
     
    def CalLineTotalPrice(self):
        if self.get_children():
            for iid in self.get_children():
                newrow = list(self.item(iid,'values')).copy()
                newrow[14] = "{:.2f}".format(round(int(newrow[10])* float(newrow[12]),2))
                self.item(iid, values=tuple(newrow))
        else:
             messagebox.showerror('Calculate All Lines' ,'No Parts in RFQ',parent = self.root)

        
class EntryPopup(tk.Entry):
    
    

    def __init__(self, parent, iid, text, **kw):
        ''' If relwidth is set, then width is ignored '''
        super().__init__(parent, **kw)
        self.tv = parent
        self.iid = iid
        

        self.insert(0, text) 
        self.selection_range(0, 'end')
        # self['state'] = 'readonly'
        # self['readonlybackground'] = 'white'
        # self['selectbackground'] = '#1BA1E2'
        self['exportselection'] = False
    
        self.focus_force()
        self.bind("<Return>", self.on_return)
        self.bind("<Control-a>", self.select_all)
        self.bind("<Escape>", lambda *ignore: self.destroy())
        self.bind("<FocusOut>", self.onClickOutsideEntry)
        self.tv.bind("<MouseWheel>",self.onClickOutsideEntry )
        

    def on_return(self, event):
        
        try:
            assert type(float(self.get())) == float
            
        except ValueError:
            if self.get() == 'None' or self.get() == '':
                messagebox.showinfo('Edit Unit Price', 'Do you wish to keep \nUnit Price as None?')
            else :
                messagebox.showerror('Edit Unit Price', 'Invalid Entry for Unit Price')
            

        newrow = self.tv.item(self.iid)['values'].copy()
        newrow[12] = self.get()
        newrow[14] = "{:.2f}".format(round(int(newrow[10])* float(newrow[12]),2))
        self.tv.item(self.iid, values=newrow)
        self.destroy()

    def select_all(self, *ignore):
        ''' Set selection on the whole text '''
        self.selection_range(0, 'end')

        # returns 'break' to interrupt default key-bindings
        return 'break'
    

    def onClickOutsideEntry(self,event):
        self.destroy()



    















