from winreg import *
import windows_tools.registry
import windows_tools.wmi_queries


# узнаем ключ Windows
def decode_key(rpk):
    rpk_offset = 52
    i = 24
    sz_possible_chars = "BCDFGHJKMPQRTVWXY2346789"
    key = ""

    rpk = list(rpk)

    is_win8 = int((rpk[66] / 6) and 1)
    rpk[66] = (rpk[66] + 247) - 255 if (rpk[66] + 247) >= 255 else (rpk[66] + 247) or (is_win8 + 2) * 4
    # rpk[66] = 0

    while i >= 0:
        current = 0
        j = 14
        while j >= 0:
            current = current * 256
            current = rpk[j + rpk_offset] + current
            rpk[j + rpk_offset] = (int(current / 24) if int(current / 24) <= 255 else 255)
            current = current % 24
            j = j - 1
            last = current
        i = i - 1
        key = sz_possible_chars[current] + key

    keypart1 = key[1:last + 1]
    keypart2 = key[last + 1:len(key)]
    key = f'{keypart1}N{keypart2}'

    key = f'{key[0:5]}-{key[5:10]}-{key[10:15]}-{key[15:20]}-{key[20:25]}'
    return key


def get_key_from_reg(key='SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion', value='DigitalProductID'):
    arch_keys = [0, KEY_WOW64_32KEY, KEY_WOW64_64KEY]
    for arch in arch_keys:
        try:
            key = OpenKey(HKEY_LOCAL_MACHINE, key, 0, KEY_READ | arch)
            value, type = QueryValueEx(key, value)
            return decode_key(list(value))
        except (FileNotFoundError, TypeError) as e:
            pass


def get_key_from_wmi():
    """
    Searches WMI for productkey
    """

    product_key = windows_tools.wmi_queries.query_wmi(
        "SELECT OA3xOriginalProductKey FROM SoftwareLicensingService",
        name="windows_tools.product_key.get_windows_product_key_from_wmi",
    )
    try:
        return product_key[0]["OA3xOriginalProductKey"]
    except (TypeError, IndexError, KeyError, AttributeError):
        return None


windows_key = get_key_from_wmi()

windows_key2 = 'WMI: ' + windows_key if windows_key != '' and windows_key != None else 'REG: ' + get_key_from_reg()
print(windows_key2)
input('Press eny key to close...')
