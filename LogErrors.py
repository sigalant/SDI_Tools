import sys
import os.path
import os
import traceback
import datetime

def handle_exception(exc_type, exc_value,exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        print("Nice CTRL-C, Nerd!")
        return
    filename, line, dummy, dummy = traceback.extract_tb(exc_traceback).pop()
    filename = os.path.basename(filename)
    error = "%s: %s" % (exc_type.__name__, exc_value)
    print("Closed due to an error. This is the full error report:")
    print()
    print("".join(traceback.format_exception(exc_type,exc_value,exc_traceback)))
    if not os.path.exists('./Error Log'):
        os.makedirs('./Error Log')
    filename = "./Error Log/"+str(datetime.datetime.now()) + ".txt"
    f = open(filename.replace(':','.'),'x')
    f.write("".join(traceback.format_exception(exc_type,exc_value,exc_traceback)))
    f.write(error)
    f.close()
    sys.exit(1)

sys.excepthook = handle_exception
