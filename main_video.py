import cv2
import os
from datetime import datetime
from simple_facerec import SimpleFacerec
import openpyxl

# Encode faces from a folder
sfr = SimpleFacerec()
#sfr.load_encoding_images("images/")
length = sfr.load_encoding_images("images/")
# Load Camera
cap = cv2.VideoCapture(0)
list_detail=[]
list_name=[]
list_info=["Ma sinh Vien","Thoi gian"]
input_detail=[]
input_detail.append(list_info)



def output_Excel(input_detail,output_excel_path):
  #Xác định số hàng và cột lớn nhất trong file excel cần tạo
  row = len(input_detail)
  column = len(input_detail[0])

  #Tạo một workbook mới và active nó
  wb = openpyxl.Workbook()
  ws = wb.active
  
  #Dùng vòng lặp for để ghi nội dung từ input_detail vào file Excel
  for i in range(0,row):
    for j in range(0,column):
      v=input_detail[i][j]
      ws.cell(column=j+1, row=i+1, value=v)

  #Lưu lại file Excel
  wb.save(output_excel_path)


while True:

    ret, frame = cap.read()
    if not ret:
        break
    # Detect Faces
    face_locations, face_names = sfr.detect_known_faces(frame)
    for face_loc, name in zip(face_locations, face_names):
        y1, x2, y2, x1 = face_loc[0], face_loc[1], face_loc[2], face_loc[3]

        cv2.putText(frame, name,(x1, y1-10), cv2.FONT_HERSHEY_DUPLEX, 1, (102, 102, 255), 2)
        cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 255, 255), 2)
        if len(list_name) < int(length):
            if name not in list_name and name != "Unknown":
                now = datetime.now()
                dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            
                #print("date and time =", dt_string)
                list_name.append(name)
                list_detail.append(name)
                list_detail.append(dt_string)
                input_detail.append(list_detail)
                print(input_detail)
                list_detail=[]
                print(list_detail)
    cv2.imshow("Frame", frame)
    output_excel_path='./DiemDanh.xlsx'
    if os.path.isfile(output_excel_path):
        os.remove("DiemDanh.xlsx")
        open("DiemDanh.xlsx", "x")
        output_Excel(input_detail,output_excel_path)
    else:
        open("DiemDanh.xlsx", "x")
        output_Excel(input_detail,output_excel_path)
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()