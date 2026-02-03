"""
Image Encoder - Steganography Tool
Encodes the C2 server IP address into the red channel of the first 4 pixels of an image.

Example:
    IP: 192.168.174.131
    Original pixels: (0,0,0,0), (0,0,0,0), (0,0,0,0), (0,0,0,0)
    Encoded pixels:  (192,0,0,0), (168,0,0,0), (174,0,0,0), (131,0,0,0)
"""

from PIL import Image
import numpy as np
import argparse
import os

def encode_ip(ip_address: str, input_path: str, output_path: str) -> None:
    """
    Encode an IP address into the first 4 pixels of an image.
    
    Args:
        ip_address: The IP address to encode (e.g., "192.168.174.131")
        input_path: Path to the input image
        output_path: Path to save the encoded image
    """
    # Read the image
    img = Image.open(input_path)
    pixels = np.array(img)
    
    # Convert IP to list of integers
    ip_bytes = list(map(int, ip_address.split('.')))
    
    if len(ip_bytes) != 4:
        raise ValueError("Invalid IP address format. Expected format: x.x.x.x")
    
    for byte in ip_bytes:
        if byte < 0 or byte > 255:
            raise ValueError(f"Invalid IP byte value: {byte}. Must be 0-255")
    
    # Modify the first 4 pixels to store the IP in the red channel
    for i in range(4):
        pixels[0, i, 0] = ip_bytes[i]  # Store IP byte in red channel
    
    # Save the modified image
    encoded_img = Image.fromarray(pixels)
    encoded_img.save(output_path)
    
    print(f"[+] IP address '{ip_address}' encoded successfully!")
    print(f"[+] Encoded image saved to: {output_path}")


def decode_ip(image_path: str) -> str:
    """
    Decode an IP address from the first 4 pixels of an image.
    
    Args:
        image_path: Path to the encoded image
        
    Returns:
        The decoded IP address as a string
    """
    img = Image.open(image_path)
    pixels = np.array(img)
    
    # Extract IP from red channel of first 4 pixels
    ip_bytes = [pixels[0, i, 0] for i in range(4)]
    ip_address = '.'.join(map(str, ip_bytes))
    
    return ip_address


def main():
    parser = argparse.ArgumentParser(
        description="Encode/Decode IP address in image using steganography"
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Commands')
    
    # Encode command
    encode_parser = subparsers.add_parser('encode', help='Encode IP into image')
    encode_parser.add_argument('--ip', required=True, help='IP address to encode')
    encode_parser.add_argument('--input', required=True, help='Input image path')
    encode_parser.add_argument('--output', required=True, help='Output image path')
    
    # Decode command
    decode_parser = subparsers.add_parser('decode', help='Decode IP from image')
    decode_parser.add_argument('--input', required=True, help='Encoded image path')
    
    args = parser.parse_args()
    
    if args.command == 'encode':
        encode_ip(args.ip, args.input, args.output)
    elif args.command == 'decode':
        ip = decode_ip(args.input)
        print(f"[+] Decoded IP address: {ip}")
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
