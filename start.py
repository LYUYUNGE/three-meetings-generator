import os

import server

server.HOST = "0.0.0.0"
server.PORT = int(os.environ.get("PORT", "8765"))
server.main()
