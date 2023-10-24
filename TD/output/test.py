#python 

'''
setup instruction 
for Install pythonosc
pip install python-osc
'''

from pythonosc import dispatcher
from pythonosc import osc_server
from pythonosc import udp_client
import win32print
import win32ui
from PIL import Image, ImageWin


def messege_handler(unused_addr, *p):
    try:
        print(p)
        client.send_message("/message_echo", p)
        #
        # Constants for GetDeviceCaps
        #
        #
        # HORZRES / VERTRES = printable area
        #
        HORZRES = 8
        VERTRES = 10
        #
        # LOGPIXELS = dots per inch
        #
        LOGPIXELSX = 88
        LOGPIXELSY = 90
        #
        # PHYSICALWIDTH/HEIGHT = total area
        #
        PHYSICALWIDTH = 110
        PHYSICALHEIGHT = 111
        #
        # PHYSICALOFFSETX/Y = left / top margin
        #
        PHYSICALOFFSETX = 112
        PHYSICALOFFSETY = 113

        printer_name = win32print.GetDefaultPrinter ()
        print(printer_name)
        file_name = "Score.jpg"
        #devmode = printer_name['pDevMode']
        #devmode.PaperLength = 381  # in 0.1mm
        #devmode.PaperWidth = 381
        #win32print.SetPrinter(printer_name, 0, d, 0)

        #
        # You can only write a Device-independent bitmap
        #  directly to a Windows device context; therefore
        #  we need (for ease) to use the Python Imaging
        #  Library to manipulate the image.
        #
        # Create a device context from a named printer
        #  and assess the printable size of the paper.
        #
        hDC = win32ui.CreateDC ()
        hDC.CreatePrinterDC (printer_name)
        printable_area = hDC.GetDeviceCaps (HORZRES), hDC.GetDeviceCaps (VERTRES)
        printer_size = hDC.GetDeviceCaps (PHYSICALWIDTH), hDC.GetDeviceCaps (PHYSICALHEIGHT)
        printer_margins = hDC.GetDeviceCaps (PHYSICALOFFSETX), hDC.GetDeviceCaps (PHYSICALOFFSETY)

        #
        # Open the image, rotate it if it's wider than
        #  it is high, and work out how much to multiply
        #  each pixel by to get it as big as possible on
        #  the page without distorting.
        #
        bmp = Image.open (file_name)
        if bmp.size[0] > bmp.size[1]:
          bmp=bmp.transpose(0)

        ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
        scale = min (ratios)
        print(scale)
        #
        # Start the print job, and draw the bitmap to
        #  the printer device at the scaled size.
        
        hDC.StartDoc (file_name)
        hDC.StartPage ()

        dib = ImageWin.Dib (bmp)
        scaled_width, scaled_height = [int (scale * i) for i in bmp.size]
        x1 = int ((printer_size[0] - scaled_width) / 2)
        y1 = int ((printer_size[1] - scaled_height) / 2)
        print("original:" ,x1, y1)
        y1 = int ((printer_size[1] - scaled_height) / 2) - 11400
        x2 = x1 + scaled_width
        y2 = y1 + scaled_height
        print ("print rect: ", x1, y1, x2, y2)
        dib.draw (hDC.GetHandleOutput (), (x1, y1, x2, y2))

        hDC.EndPage ()
        hDC.EndDoc ()
        hDC.DeleteDC ()
    except ValueError: pass
    print("received")

dispatcher = dispatcher.Dispatcher()
dispatcher.map("/message", messege_handler)
server = osc_server.ThreadingOSCUDPServer(("127.0.0.1", 10000), dispatcher)
print("Serving on {}".format(server.server_address))
    
client = udp_client.SimpleUDPClient("127.0.0.1", 20000)

server.serve_forever()
