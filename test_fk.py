import cv2
from pyzbar import pyzbar
import json
import pandas as pd
from datetime import datetime
import serial

barcode_data = {}

def load_barcode_data():
    try:
        with open('barcode_data.json', 'r') as file:
            barcode_data.update(json.load(file))
    except FileNotFoundError:
        pass

def save_barcode_data():
    with open('barcode_data.json', 'w') as file:
        json.dump(barcode_data, file)

def send_to_arduino(number):
    # Define the serial port and baud rate
    port = 'COM9'  # Replace 'COM3' with the appropriate serial port
    baud_rate = 9600

    # Initialize the serial connection
    arduino = serial.Serial(port, baud_rate)
    
    # Wait for a moment to establish the connection
    arduino.timeout = 2
    arduino.write(str(number).encode())

    # Close the serial connection
    arduino.close()

def detect_barcode():
    cap = cv2.VideoCapture(0)

    while True:
        _, frame = cap.read()

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        barcodes = pyzbar.decode(gray, symbols=[pyzbar.ZBarSymbol.CODE39])

        if len(barcodes) > 0:
            print("Barcode Detected!")
            for barcode in barcodes:
                barcode_data['number'] = barcode.data.decode('utf-8')

                if barcode_data['number'] in barcode_data:
                    print("Barcode Data:")
                    print("Number:", barcode_data[barcode_data['number']]['number'])
                    print("Name:", barcode_data[barcode_data['number']]['name'])
                else:
                    barcode_data[barcode_data['number']] = {}
                    barcode_data[barcode_data['number']]['number'] = barcode_data['number']
                    barcode_data[barcode_data['number']]['name'] = input("Enter the name for the barcode: ")

                    print("Barcode Data:")
                    print("Number:", barcode_data['number'])
                    print("Name:", barcode_data[barcode_data['number']]['name'])
                    
                barcode_data[barcode_data['number']]['sem'] = input("Enter the semester: ")
                barcode_data[barcode_data['number']]['intime'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Convert barcode_data to DataFrame
                df = pd.DataFrame(barcode_data).T

                # Save DataFrame to Excel
                df.to_excel('barcode_data.xlsx', index=False)
                
                # Send the number to Arduino
                number = int(barcode_data['number'])
                send_to_arduino(number)
        else:
            print("No Barcode Detected!")

        cv2.imshow("Barcode Scanner", frame)

        save_barcode_data()

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

load_barcode_data()
detect_barcode()