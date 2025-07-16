    #!/usr/bin/env bash

    # Install pandoc and any dependencies
    apt-get update && apt-get install -y pandoc

    # Install Python dependencies
    pip install -r requirements.txt
