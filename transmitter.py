import socket
import excel_tool

def send_file(filename, host, port):
    with open(filename, 'rb') as f:
        data = f.read()
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((host, port))
        s.sendall(data)
    print('File sent successfully.')

if __name__ == '__main__':
    excel_tool.serialize_excel(True, '221综测表.xlsx', 'temp_data.txt')
    send_file('temp_data.txt', '192.168.1.4', 9120)