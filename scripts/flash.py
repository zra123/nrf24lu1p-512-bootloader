#!/bin/env python
import usb.core
import usb.util
import usb.backend.libusb1
import libusb_package
import time
import bootloader
import sys
import intelhex
import os
import click
from typing import Optional
from ctypes import wintypes, windll, create_unicode_buffer

backend = usb.backend.libusb1.get_backend(find_library=libusb_package.find_library)

ID_VENDOR = 0x1915
ID_PRODUCT = 0x0101

# TODO: rewrite this with argparse #
# TODO: use the fact that we know endpoint values to clean this up #

debug = True

# def usb_setup():
# dev = usb.core.find()
dev = usb.core.find(idVendor=ID_VENDOR, idProduct=ID_PRODUCT, backend=backend)
if not dev:
    print("Unable to find dongle")
    input("Press Enter to exit...\n")
    sys.exit()

cfg = dev.get_active_configuration()
intf = cfg[(0,0)]

ep1in = usb.util.find_descriptor(
    intf,
    # match the first OUT endpoint
    custom_match = \
    lambda e: \
        e.bEndpointAddress == 0x81)
ep1out = usb.util.find_descriptor(
    intf,
    # match the first OUT endpoint
    custom_match = \
    lambda e: \
        e.bEndpointAddress == 0x01)


def getForegroundWindowTitle() -> Optional[str]:
    hWnd = windll.user32.GetForegroundWindow()
    length = windll.user32.GetWindowTextLengthW(hWnd)
    buf = create_unicode_buffer(length + 1)
    windll.user32.GetWindowTextW(hWnd, buf, length + 1)
    
    # 1-liner alternative: return buf.value if buf.value else None
    if buf.value:
        return buf.value
    else:
        return None

def bytes_to_str(byte_array):
    return ''.join(["{:02x}".format(byte) for byte in byte_array])

def usb_cmd(cmd, arg=None):
    raw_cmd = bytes([cmd])
    if arg != None:
        raw_cmd = bytes([cmd, arg])
    try:
        ep1out.write(raw_cmd)
    except Exception as e:
        print("Error: Please install the driver from Zadig")
        input("Press Enter to exit...\n")
        sys.exit()

def read_usb_in():
    data = ep1in.read(64, 1000)
    return data

def print_usb_response():
    print(bytes_to_str(usb_get_response()))

def usb_get_response():
    # todo error handling
    return ep1in.read(64, 1000)

def stp_on():
    usb_cmd(bootloader.CMD_STP_ON)
#    print_usb_response()

def bootloader_version():
    usb_cmd(bootloader.CMD_VERSION)
    print_usb_response()

def flash_read_disable():
    usb_cmd(bootloader.CMD_READ_DISABLE)
    print_usb_response()

def flash_read_block(block):
    usb_cmd(bootloader.CMD_READ_FLASH, block)
    return usb_get_response()

def flash_erase_page(page_num):
    usb_cmd(bootloader.CMD_ERASE_PAGE, page_num)
    print_usb_response()

def flash_select_half(half):
    usb_cmd(bootloader.CMD_SELECT_FLASH, half)
    return usb_get_response()

def mcu_reset():
    usb_cmd(bootloader.CMD_RESET)

def flash_write_page(page_num, page):
    usb_cmd(bootloader.CMD_WRITE_INIT, page_num)

    for a,b in [(i,i+64) for i in range(0, 512, 64)]:
        block = page[a:b]
        if len(block) < 64:
            block += bytes([0xff] * (64 - len(block)))
        ep1out.write(block)
        ep1in.read(64, 10000)

def is_empty_flash_data(data):
    for b in data:
        if b != 0xff:
            return False
    return True

def chunk_iterator(data, chunk_size):
    for i in range(0, len(data), chunk_size):
        yield data[i:i + chunk_size]

def flash_read_to_hex(size=16):
    # result = []
    if size not in [1, 16, 32]:
        raise Exception("Cannot read flash size {} only 1 or 16 or 32".format(size))

    page_size = 0x200
    block_size = 0x40
    blocks_per_16kb = 0x100
    hexfile = intelhex.IntelHex()
    if size == 1:
        cur_addr = 0x7e00
        blocks_per_size = 0xf8
    else:
        cur_addr = 0x0000
        blocks_per_size = 0x0

    def read_16kb_region():
        chunks_per_block = 4
        chunk_size = block_size // chunks_per_block
        nonlocal cur_addr
        for i in range(blocks_per_size, blocks_per_16kb):
            block = flash_read_block(i)
            # break the block into chunks to we can check which regions
            # actuall have flash data
            for chunk in chunk_iterator(block, chunk_size):
                if not is_empty_flash_data(chunk):
                    hexfile.puts(cur_addr, bytes(chunk))
                cur_addr += chunk_size

    if size == 16 or size == 32:
        flash_select_half(0)
        read_16kb_region()
    if size == 32 or size == 1:
        flash_select_half(1)
        read_16kb_region()
    return hexfile

def stp_off():
        hexfile_page = flash_read_to_hex(size=1)
        data_page = hexfile_page.tobinarray()
        #print(hexfile_page.dump())

        if len(data_page) >= 0x1f0 and click.confirm("STP protection is ON! Disable protection?",default=True):
            for i in range(0x1f0,0x1ff):
                data_page[i] = 0xff

            flash_write_page(63, data_page)

            hexfile_page = flash_read_to_hex(size=1)
            data_page = hexfile_page.tobinarray()
            if len(data_page) <= 0x1f0:
                print("STP OFF")
            else:
                raise Exception("STP protection is still ON!")

        #mcu_reset()

def hex_dump(data):
    addr = 0
    for block in [data[i:i+64] for i in range(0, len(data), 64)]:
        print("{:04x}".format(addr), bytes_to_str(block))
        addr += 64

def help():
    print('''
    -r16,  read_16	 - Read 16kb to file
    -r32,  read_32	 - Read 32kb to file
    -v,    version	 - Print bootloader version
    -w,    write	 - Write file to dongle
    -off,  stp_off	 - Disable FSR.STP register
    -on,   stp_on	 - Enable FSR.STP register
    -rd,   read_disable	 - Turn on flash MainBlock readback disable''')

check_cmd = os.path.exists(getForegroundWindowTitle())

if check_cmd:
    arg = "write"
else:
    try:
        arg = sys.argv[1]
    except IndexError:
        help()
        sys.exit()

# TODO: cleanup cmd line handling

if arg == "read_16" or arg == "-r16":
    try:
        outfile = sys.argv[2]
    except IndexError:
        print("Out_file.hex or .bin")
        sys.exit()
    ihex = flash_read_to_hex(size=16)
    ihex.dump()
    ext = os.path.splitext(os.path.basename(outfile))[1]
    if ext == ".hex":
        ihex.tofile(outfile, "hex")
    else:
        ihex.tofile(outfile, "bin")

elif arg == "read_32" or arg == "-r32":
    try:
        outfile = sys.argv[2]
    except IndexError:
        print("Out_file.hex or .bin")
        sys.exit()
    ihex = flash_read_to_hex(size=32)
    ihex.dump()
    ext = os.path.splitext(os.path.basename(outfile))[1]
    if ext == ".hex":
        ihex.tofile(outfile, "hex")
    else:
        ihex.tofile(outfile, "bin")

elif arg == "version" or arg == "-v":
    bootloader_version()

# TODO: read back flash and verify what we wrote #
elif arg == "write" or arg == "-w":
    flash_size = 0x8000
    page_size = 0x0200
    num_pages = flash_size // page_size

    if check_cmd:
        write_file = sys._MEIPASS + '\\watchman_dongle_mod.hex'
    else:
        try:
            write_file = sys.argv[2]
        except IndexError:
            print("Write_file.hex or .bin")
            sys.exit()

    hexfile = intelhex.IntelHex()
    ext = os.path.splitext(os.path.basename(write_file))[1]
    if ext == ".hex":
        hexfile.loadhex(write_file)
    elif ext == ".bin":
        hexfile.loadbin(write_file)
    elif ext != ".hex" or ext != ".bin":
        raise TypeError("Write_file.hex or .bin")

    if hexfile.maxaddr() > flash_size:
        raise "file too large"

    hexfile.padding = 0xff
    data = hexfile.tobinarray(start=0x0000, end=flash_size)

    stp_off()

    print("Starting to write file:")

    for page_num in range(0, num_pages):
        page_start = page_num * page_size
        page_end = (page_num+1) * page_size
        page_data = data[page_start:page_end]

        is_empty_page = True
        for b in page_data:
            if b != 0xff:
                is_empty_page = False
                break

        if is_empty_page:
            continue

        # print("{:x} {:x} {:x} {}".format(page_start, page_end, len(page_data), page_data))

        flash_write_page(page_num, page_data)

    mcu_reset()
    print("Done")
    time.sleep(5)

elif arg == "stp_off" or arg == "-off":
    stp_off()
    print("Done")

elif arg == "stp_on" or arg == "-on":
    stp_on()
    print("Done")

elif arg == "read_disable" or arg == "-rd":
    flash_read_disable()

elif arg == "help" or arg == "-h":
    help()
    pass

elif arg == "new_cmd":
    pass

# # It may raise USBError if there's e.g. no kernel driver loaded at all
# if reattach:
#     print("reattach")
#     while dev.is_kernel_driver_active(intf_num):
#         pass
#     dev.attach_kernel_driver(intf_num)
