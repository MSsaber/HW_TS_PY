
import socket
import excel_tool

def receive_file(filename, host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind((host, port))
        s.listen()
        conn, addr = s.accept()
        with conn:
            with open(filename, 'wb') as f:
                while True:
                    data = conn.recv(1024)
                    if not data:
                        break
                    f.write(data)
    print('File received successfully.')

if __name__ == '__main__':
    receive_file('receive_excel.txt', '192.168.1.2', 9120)
    excel_tool.serialize_excel(False, 'receive_excel.txt', '221zc_table.xlsx')