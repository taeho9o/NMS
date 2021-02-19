import os

from pyftpdlib.authorizers import DummyAuthorizer
from pyftpdlib.handlers import FTPHandler
from pyftpdlib.servers import FTPServer

FTP_HOST = '0.0.0.0'
FTP_PORT = 21

FTP_ADMIN_DIR = os.path.join(os.getcwd(), 'storage')

if __name__ == '__main__':
    authorizer = DummyAuthorizer()

    authorizer.add_user('user', '12345', FTP_ADMIN_DIR, perm='elradfmwMT')

    handler = FTPHandler
    handler.banner = "FTP Server."

    handler.authorizer = authorizer
    handler.passive_ports = range(60000, 65535)

    address = (FTP_HOST, FTP_PORT)
    server = FTPServer(address, handler)

    server.max_cons = 256
    server.max_cons_per_ip = 5

    server.serve_forever()
