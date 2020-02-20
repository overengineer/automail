import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
import openpyxl, glob, os
from getpass import getpass
from tkinter import *

class VerticalScrolledFrame(Frame):
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(2*(event.num-4)-1, "units")
    def __init__(self, parent, *args, **kw):
        Frame.__init__(self, parent, *args, **kw)            
        # create a canvas object and a vertical scrollbar for scrolling it
        vscrollbar = Scrollbar(self, orient=VERTICAL)
        vscrollbar.pack(fill=Y, side=RIGHT, expand=FALSE)
        canvas = Canvas(self, bd=0, highlightthickness=0,
                        yscrollcommand=vscrollbar.set)
        canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        canvas.bind_all("<Button-4>", self._on_mousewheel)
        canvas.bind_all("<Button-5>", self._on_mousewheel)
        canvas.pack(side=LEFT, fill=BOTH, expand=TRUE)
        vscrollbar.config(command=canvas.yview)
        # reset the view
        canvas.xview_moveto(0)
        canvas.yview_moveto(0)
        # create a frame inside the canvas which will be scrolled with it
        self.interior = interior = Frame(canvas)
        interior_id = canvas.create_window(0, 0, window=interior,
                                           anchor=NW)
        self.canvas = canvas
        # track changes to the canvas and frame width and sync them,
        # also updating the scrollbar
        def _configure_interior(event):
            # update the scrollbars to match the size of the inner frame
            size = (interior.winfo_reqwidth(), interior.winfo_reqheight())
            canvas.config(scrollregion="0 0 %s %s" % size)
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # update the canvas's width to fit the inner frame
                canvas.config(width=interior.winfo_reqwidth())
        interior.bind('<Configure>', _configure_interior)
        def _configure_canvas(event):
            if interior.winfo_reqwidth() != canvas.winfo_width():
                # update the inner frame's width to fill the canvas
                canvas.itemconfigure(interior_id, width=canvas.winfo_width())
        canvas.bind('<Configure>', _configure_canvas)

def send_mail(server, send_from, send_to, subject, text, text_type='html', files=None):
    if not isinstance(send_to, list):
        send_to = [send_to]
    if not isinstance(files, list):
        files = [files]
    msg = MIMEMultipart('alternative')
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(text, text_type))
    for f in files or []:
        with open(f, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=basename(f)
            )
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
        msg.attach(part)
    server.sendmail(send_from, send_to, msg.as_string())

def load_receivers(path):
    wb = openpyxl.load_workbook(
            filename=path, 
            data_only=True, 
            read_only=True
        )
    rows = [row[:2] for row in list(wb.active.values)[1:]]
    return [(name.strip(), mail.strip()) for name, mail in rows if name and mail]

class Automail():
    def __init__(self, config_path):
        self.config_path = config_path
        self.load_config()

    def load_config(self):
        wb = openpyxl.load_workbook(
            filename=self.config_path, 
            data_only=True, 
            read_only=True
        )
        self.table = list(wb.active.values) + [('Password',None)]
        self.config = {key: value for key, value in self.table}

    def send_mails(self):
        if not hasattr(self, 'msg_frame'):
            self.msg_frame = VerticalScrolledFrame(self.app)
            self.msg_frame.pack()
        try:
            values = [self.fields[key].get() for key, value in self.table]
            yandex_mail,receivers_path,attachments_path,title,template_path,yandex_pass=values
            server = smtplib.SMTP_SSL('smtp.yandex.com.tr','465')
            server.ehlo()
            server.login(yandex_mail,yandex_pass)

            with open(template_path) as f:
                msg = f.read()
            for target_name, target_email in load_receivers(receivers_path):
                try:
                    if target_name and target_email:
                        attachments = glob.glob(
                            os.path.join(attachments_path,target_name+'.*'))
                        assert attachments, 'Attachment not found!'
                        send_mail(
                            server,yandex_mail,target_email,title,
                            msg.format(NAME=target_name),
                            text_type='html',
                            files=attachments
                        )
                except Exception as ex:
                    self.message('%s %s Gönderilemedi\n%s'%(
                        target_name, target_email, ex),color='red')
            self.message('İşlem tamamlandı.')
        except Exception as ex:
            self.message(str(ex),color='red')
        finally:
            server.quit()

    def gui(self):
        self.fields = {}
        for prompt, value in self.config.items():
            self.fields[prompt] = StringVar()
            label = Label(self.app, text=prompt,width=50).pack()
            show = {True:'*', False:None}[prompt == 'Password']
            entry = Entry(self.app, textvariable=self.fields[prompt], show=show,width=50)   
            entry.insert(END, (value or '').strip())
            entry.pack()
        self.submit = Button(self.app, text='Gönder',command=self.send_mails)
        self.submit.pack()  
        def return_handler(event):
            self.send_mails()
        def esc_handler(event):
            self.app.quit()
        self.app.bind('<Return>',return_handler)
        self.app.bind('<Escape>',esc_handler)
        self.app.mainloop() 

    def message(self,msg,color='green'):
        Label(self.msg_frame.interior,text=msg,
            bg='Black',fg=color,justify=LEFT,anchor='w',width=50).pack()
        print(msg)

    def __enter__(self, *args):
        self.app = Tk()
        return self
    def __exit__(self, *args):
        self.app.quit()

with Automail('Konfigurasyon.xlsx') as automail:
    automail.gui()  



