import socket

def send_file(filename, host, port):
    with open(filename, 'rb') as f:
        data = f.read()
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((host, port))
        s.sendall(data)
    print('File sent successfully.')

if __name__ == '__main__':
    send_file('temp_data.txt', 'localhost', 9120)