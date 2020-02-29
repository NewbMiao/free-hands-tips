FROM golang:1.14
WORKDIR /go/src/app
RUN apt-get update \
    && apt-get install -y gdb \
    && rm -rf /var/lib/apt/lists/*
# Enable cyrus-and/gdb-dashboard
RUN wget -O /root/.gdbinit https://git.io/.gdbinit \
    && mkdir /root/.gdbinit.d \
    # GDB Pretty printers and convenience functions for Go's runtime structures.
    && echo "add-auto-load-safe-path /usr/local/go/src/runtime/runtime-gdb.py" > /root/.gdbinit.d/init \
    && wget -O get-pip.py https://bootstrap.pypa.io/get-pip.py \
    && python get-pip.py && pip install pygments && rm get-pip.py
# Also can enable dlv
RUN go get -u github.com/go-delve/delve/cmd/dlv

# usage:
# docker build -t newbmiao/gdb-go:1.14 .
# #Bind current dir as volume to debuger
# docker run --rm -it --security-opt="apparmor=unconfined" \
#    --security-opt="seccomp=unconfined" \
#    --cap-add=SYS_PTRACE -v "$(pwd):/go/src/app" \
#    newbmiao/gdb-go:1.14 bash
# #Build go without compressdwarf to map code
# go build -ldflags=-compressdwarf=false -gcflags=all="-N -l" -o test test.go