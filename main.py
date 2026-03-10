import sys
import gui
import service

def main():

    if len(sys.argv) > 1 and sys.argv[1] == "service":
        service.run_service()
    else:
        gui.run_gui()

if __name__ == "__main__":
    main()