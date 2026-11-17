# -*- coding: utf-8 -*-
# =============================================================================
#  ANALIZ TAKIP - ONE-TIME SETUP TEST  (analiz_takip_kurulum_test_16.py)
#  Run in Ansys Mechanical (Automation > Scripting > Run Script).
#  1) writes the Analysis Tracker files (PowerShell version, package 14) to a folder
#  2) tries to create the Excel workbook (.xlsm) with the macro and 4 buttons
#  3) shows a PASS / FAIL report of every step (also saved as kurulum_test.log)
#  Nothing here changes any security setting.
# =============================================================================
import clr
import os
import base64
import zlib
import traceback
from datetime import datetime
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System import Type, Activator, Array, Byte
from System.IO import File, Directory, Path
from System.Text import UTF8Encoding
from System.Windows.Forms import (Form, TextBox, Button, FolderBrowserDialog, DialogResult, DockStyle,
                                  ScrollBars, FormStartPosition, MessageBox, FlowLayoutPanel, FlowDirection)
from System.Drawing import Font, Size

# Target folder. Leave empty to be asked.  Example: HEDEF = r"C:\AnalizTakip"
HEDEF = r""
EXCEL_ADI = "Analysis Tracker.xlsm"

PAYLOAD = {
    'AnalizTakip/AdminAccounts.txt': (
        "eNpVjrGKwmAQhHvBdxhIcweSB7ATz+JATXNgY7Mkv8nC/++G3Q053/5yprKbYfhmpkIzBqtQ3oO6jle9yMLCHkah5jUaSbixdDo7"
        "qG11ksCYDJkl7cCCGBIeagVfzeXwfb1PvqQf6ZfKmNMeh+PldFcvJJ/bTYXzgjk8yIKlx8wxoAJZAveilroaP/+FbP7acRXMg66E"
        "v8ZG096ogFf7dhg0hRYKbinnZ73d/AGwgFGs"
    ),
    'AnalizTakip/AnalizTakip.ps1': (
        "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"
        "L9MCgVlV8OkqJS7R1H0oxHwLmRuxz/iooxvcLpRo/JahqG9Q8ObyffRO7YoX07kmGrtzV96xXPWS27Oj+gqPVPlk12a9CD2lma0K"
        "WrdzDMeJfIDHZ56V/Aa9aiasrX/GmW6dHWb50DvhT+dXMEVOy8fOO/cbJcXLOBXCueLajcTMSW9jrcNVJ3YzEpTCyk/aC1JYWSTX"
        "F5KWDl24cytu6C3rUK1VHdjANY/LQudA1i6K7FRi97wIm58NK8U0mbTr6o7qyevyWjlzP5gVvtx8X3ijYysJ6neIaaL1VGWv/Mm0"
        "dXNzcoSQBSLmvsGIfRbMcs6JkrDWiePUUrOSNvyzqsJSaunDjQWnYk8fUwQFqteLx9RPFz74cZGloIThswRXIcQavfGiSe+Ht22Z"
        "1VL5DLhCBgc3G2CGLB5JKsOJnxXrrF6xiGMssJwbEqnE8Sm/E4K34g6emxG2GaFK9T7ojVYFt4sukfVKsu1M2TIA6k4WpWbypujk"
        "KCyRfs6mMJvvxAQzZ54zvbrdidOV0hroUBltkx8DD++iWQcTwIYG5pw/2JhR7I1ui+1NLceM4XNeidkSf0zq43hYQ/1YZ+77Xqev"
        "YoSucb30o89odZ2qfTmotXfjTC1fiL/8gvn5T5zkwYS2yIG7MluZZ0BP1NFCGUOMoJdKPWFPz5Jx30Y+L3ducaCJQEI/uN6X6nFv"
        "Mu1vONefCsj/t54BFC/yAcKLuwHElQvfHzpPrskDACU2ZUO7vLemy/Cz4HaCDUDgCWWMdnRw8vzs5UGz7jnwNu3eOS4/toUyaXBc"
        "fpDYmK74CcY4uqBHJYL/h0vyl4SyMfrA8Uz6zM1FyMdB+0x/ONp/j0moqVR69fqc/Tzh7ZmLpyh7jppOm1NTn1ww6aZACFJ8AAhh"
        "e9CJmNRBG0zxA4vkNzHVHHIM7kfNmi7ZKB7y3PT4yKoA2rVzry5XgHNiFlNefZfkCnYhR3oFlejsssg/LECDGMuXPVaTvIZn+/qk"
        "fqYeMc0y6DD9GRJw/jZEW946URlW6ats5Oh35XNpEqG8mhtJi/d0YBcBQIvEq+1gthunJ1dA35giwIvPKFuI+8Q0L8EhhaSlvJvV"
        "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"
    ),
    'AnalizTakip/README.txt': (
        "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"
        "zv/3vQPn9f4AIuPOKA=="
    ),
    'AnalizTakip/Veri.ps1': (
        "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"
    ),
    'Analysis Tracker.cmd': (
        "eNpNjbFqwzAYhHeD3+EwFNJBJn2DpJASCIRADF26COmMhVX9QlKI3aHPXjVTto/j7rsdzSSQcWybxG/sg/ZrdhlD0mZmgte3YKYK"
        "myA4LIYegbS0rz1OZESZant0njVfCorUhA+P+xn07CJG8Zapb5tcdCroOkS5M+WJ3vdcCHWWS5KHQx0WmltxEi7inVnxvkadM9R1"
        "2EN9umDlfi1rbR6dtQxQH/+z7uXXxu3T6dcT9zG/dW3zB2a6UTw="
    ),
    'KURULUM.txt': (
        "eNqtVk1v20YQvRvwfxjwYrFxGLhNLzIagLHpRJBkG5Kd1kEvI3ItL7hcCkuua/qv9dCf1L/QN/yy4gY9VRdLonfmzXtv3urvP/+K"
        "L+PF7CvdxPPZNZ28p7c0v13dLm6XNMl1pStl6PqMHlXDVLqac8oNV6U7JqucahTlZcOuYquf8ZwzppQNzrnw8OCX/+11eHB4cO3K"
        "reOCsrJq2LDTdJmskrsECKwHgm/bTylXNtO00WarKwZGDYjHVHDFvqr9MW2U2SqjXDcbb/vBfHR4MDM5ZcqqQjWaXlg4yhQ1vGPU"
        "1K6Seqc9J8BjmbYqLVROOtV2KAZwO0Fb4J8aVStndKZdJPOchHS9uvq0ipc0X8Trq9Xt4QER+P+s3APbrWB3OOTQNYgtxnq+4Vzv"
        "AuLMdA+7JlQamch5S5PSWZVP6Wz6+96JMOorr/0eexvfn281xOFHJZWFqIwtBdwWoFoqwBcBlZZUjqmpwvx22taU116naFedHNMX"
        "5XT3bpXE58skqp/qY4qzQts4TUtv60q+6kENumqQzilKVzVTIDUwqakHJq3vFR2m7Vj8MaTkt7NkQedX67t4PYNrYYfCA60qQFOm"
        "ttqEfavkKVVGVNyUVWeUgqFvDUVp8hEa5vRH6fJNWeYhsNAHutBG4c+aQU5c4V3QvuWK6manAhK0ClaAW3o+grYJLTl15dvE8sao"
        "jH7tq9Lkh+jJVEUYUKXSvnDUn5xclq5gM1R4deapO9NaT8HGuStpHt+dJ+ezxTL5GoV9lSurzp1+VO/WD+zUdaltfQSFuclKA/ND"
        "PSGz4+W0r8Ou1jkMIRzzxhtSmXKDaWJTv7k4OaHJF115wPuIIiklma5LF2KEQlmP8WlmK+VqfLEsM29U9C8Htb1E+KBdkQxmohvk"
        "TKFqmKpflEbvZMdAq4YRBwyzwmB0Cq79xqD52m9oDgsYX9zxbhJ2BTsf7+jiZwy8AUo7JWz+PYuzM78t2m0v67JgGVd85Mcpz2pn"
        "3qxJ5BauVIYdnsBmU/ovPY+p17P14k8hzRZzuplhn5dxXzi42oEc2ZEGOUI3jtNcuaADVAnqFmo07kHXr91TyRgLOqysJPQx3iJJ"
        "nivutd7LgEqyh9AqRSxD5XaTgH5vOfc2adIJ8Hp1wQnbsLXZSP0aHwp6AGUOD7U7bbtm/MCtk+Cp54jmfdqJq0S5MQaDNUhP63G4"
        "i9Jke9PTI7shD9+HdHsXr2aLeDVQtxTHIJFQ9xExlSGJA/KS5Tia6lwuHkiMMEv9xue1xNYMBGiEZ1Ijq4wSNGiHIB8WdNLm/bnw"
        "KwuoUgR9Ls74QJ/8I7RaKperZz18hiUszSVUgOSIO2XsS8SDTjmPBmPOdhI2Ii5Ixt75phSorzfitSuitMiCF+m/L94wRwuCe8+n"
        "+r4muFoWyJ72QVehFac8evwzI1lhe4HyvdLtRu5bAqpzZMrtCAkmxAPb6nV5dTOlj9rJ7t/zs+H2pqTcG7kMU87RRVJ8vAr7q3K7"
        "10+4QbtTuYyAVkhl11/Vhwe73jTjL43xnm+GPvAOTcYbpDW1haZ1re1W4rpdvkzmvm+dR++6ewiPzuSSVd88lO3P1OhifGyvesnW"
        "fwCb5w7m"
    ),
    'makro.txt': (
        "eNq1WN1u47gVvg8w70Boi42EddQkM7NYZJstHMfJBpM/xN4NFhigoCXaZi2LBknF8WLfrBd9pL5Cv0NKsmTHmbSd5sKRqMNzDs/v"
        "d/ivf/zzbmGlyln/eZHJRNp3e+/29tndQuSG2alg3ZxnKyMNG2qezIRmC60mms+ZzJm0hqllzpYyT9Uypo1DbKkoxjIThoXEQf4+"
        "5DO5iBfmqMN+FVrSE4vjOGIJz9lIELvu7W8sV3rOMzZWWSr0CXFcqUI7Kfc9pjTjzEy5FinLhV0qPStJY0iGkiSSpQpic2XZlD8J"
        "ZhWxz8WzpUccaV4rOpbaYFXORYfJsTtufTrPYayKPG1sxnL/ORGZE9SBMl46MVxIZx5uZoYtp0ILxy9onD4oqcH8R/eR52bpXpkW"
        "czEfYVMakwPutXziVrCLIk+cdy5FInQmwxnrGjawWuaTiB7PlMoEz9/tMfyVROyUXfDMCL94NWbXIg9nEZYP6dTkamlr1muqHjci"
        "vBZjG8467ENEG4KptVB71661wPBcaij3LQs+b7g7iNhffmJBEL3b68OUawYvnPJTxo3SA5GE0fqcXtS5RDyltHo3+rtIbIeZTZJ6"
        "Nyke+LWBsLTtlHUXFN6cxMQX8N255JmahB8iTzdO46G0iB1sHYgMAurgL89S+S6k9fLZTrllicotl8iWjYNHQW1acB9M1ZIMcXC0"
        "05oGwonSiRfplRVzEx5FNZcHOZna0HTYkXPN59IvtM15DV/I0yZiB6yxrQods+kdeKZmsP2x3n+LNKh5+C3+G/3dmMmZem5Fucvs"
        "JW/mj8xPAgh4GvX09RgP0PRphDzK+Nx5pOMZNMtMsJbxgqkokq7GpYuR5HCzRRy0s424rrim/z408GS2Q8W8ITDvfVXw24qXo3O2"
        "ubhPjtpROlio8mxFRQdnwxdeFb5UmhlbcDstHTiDhlTZHlHqRkrN4nt82+2suk5sOqp9AjCdvdnA++w4Yo2oT6bKIGwEhyAqfZWa"
        "l8KWfgi/6AdXD/6vOr+PqBT7ty1GjULzgveLEUoFGxSjku5cTKSxKHDRK96ebbJtF9+fquJbZcx91SS9UY2gMGmnyYzS5CofU2Ts"
        "ThNSH7puaF6yl91kh9YddEUrWwszNS/sS+faDP/XO4tThgi8ALjtxb5QsncyUc4WCq3QTEWWxeJZsINbBakuVQ76zyIpyAD3Csdb"
        "sbPVghvDDgbDLjt4dNhjYFeg/FmmKbQ4oAKPECNjeh2+pbdS4h3U1BpI4lINFSCC5X69pwVy3neXMHgcJFouLOo2NAqi+KHIva4d"
        "dthp9tf1kYnViV+snNxEQ4kqstQVRcARY7lGjT9hpCPUic+FcQKdmykCpEW7yv4TlxOsKl200+v/s0+dPVgggBaVBuwiX3lLey+T"
        "EninM9y6inahksLsUvlnmOyyyLNiUrwSqOlX1D99MR7JdTHwwLqQEp5JXZ8NdrS8ocN4Dicy4QIKDFhYv03RAkcCCmiRKJ2KNIqD"
        "NyV0q8TVirdanncCYkkseNrwQVpG+mvm32cHX+/PzQq3/YPh1U2fDfrDX+5PmEaquHan8gSN7uIj+r+RqfCLc55o1SEk383sdxc/"
        "oDJ+KnSRFfPf+CJy0PwKKVLILPUDiAfpqAyokJzAhKV1LQSIrFUAXrywigyKfMlW8Vc/YbMf1Jo2g3XpoCh1aKflGnkuzWb3ronW"
        "yG5p4i56DwGONUtJHK9VFfAXMBfBbNCe+UPHPSAry1DBjpAaYgFguY6cNVkoI5SWTFS8bwmOyFrwLZ87yLsjCkHRQ6ThF0NG2emd"
        "rsJX3fhcmkXGV5dappnMhWmPHrRdwVxQI+iihvqXR5kCwZyy99tEZ9tEH2qqB55PBGiOQfQrz4ovKN4kvwBEjwfyd9pyfLib5Ay9"
        "mDymC7GbCArCG6fs4fIsfI9p9uP3QOSHh9HWjvdNRXsIoZmLZx+1KA0ZJgIgQ4VRuz16bg3WwYu8t7Q5OkJrqn5KfRAHKu/P0A+X"
        "QN3B2UfUhg+gCGjC3xruCZmtkQPe1qbYZPQDvh5/XzE6RwFFlBKUIR6tVhS0+uUmn6PDFiPfmq9RirHSag9fYHNcsSnHtwpfrZVq"
        "g7k2u7Vpu5hYyxms3cuR0AWmQJlLJHAau7mjLFDQAsU+fs7MnIU9q7PvBpGrVYURDZ+bNzaAZq8sJxEqPvWRw42C02HTIkEvarTM"
        "VTEzYpZ5fHeuUMFwWzEHFsqbVHM+06q58FToSdG4V2gUudF69vaLjxIZWtvNaRCtSxCVv1G7YnXTNIxpUu2weKgW+HVZ3lA2Wrc6"
        "4l6iyPgu7/ox7NRrXK33+KJappNVy81sv7oah+5QyIljytOoRVUmvKNo2H3v67YQum3qD4bsptt7uBuwMJV8kisEYWKiE9wVcT8C"
        "ZigIuNdSZfe87/31K6vRaGRDYewN2TKMWkF+9+nEd2jj2rgsdVlW/esNAcwseAeVMfc2pd4TzncAphRNwZU2g3DcfDGtL+7TVW7F"
        "ROgazm2Nx/5GB18cHPq8MH8jlWL7DLVaE8CDMAUaIDVFfy0gAavGO6aENw4JQTL3iOzPCXtlpump+ZwqRBD8qR5veghJqILqYwxe"
        "Bxai4mukV8En4kYBQf3B7grrx5uDfp6olC49uEmkZPuE/2hqDPYdADz0xbuJZ8cNPLsFZ9d+wR1Eym7vhtSQ0gIQjgPCmiKz5c3F"
        "KFPwc0oQLslwFQBINp8LBLUV2aoEua9f7pQhsgl0GzjXu5B8eAGkRwf2XqEWMXZw6Cpf+GH1m7GnvgYAKVexhMjxDiMNa5oXDqt5"
        "HrOWkU/dUGaak3gQXhRZVlGBotmv6QIbyfIjg/+M1bgLFGmDFGcDBbB69PbkaVnFzZUt9V1k/jez40Zu/hsePBub"
    ),
}

LOG = []


def log(ok, metin):
    LOG.append(('PASS  ' if ok else 'FAIL  ') + metin)


def bilgi(metin):
    LOG.append('      ' + metin)


def yaz(yol, b64):
    veri = zlib.decompress(base64.b64decode(b64))
    d = Path.GetDirectoryName(yol)
    if not Directory.Exists(d):
        Directory.CreateDirectory(d)
    File.WriteAllBytes(yol, Array[Byte](bytearray(veri)))
    return len(veri)


def hedef_sec():
    if HEDEF:
        return HEDEF
    fb = FolderBrowserDialog()
    fb.Description = 'Select the folder where the Analysis Tracker files and the Excel file will be created'
    if str(fb.ShowDialog()) != 'OK':
        return None
    return fb.SelectedPath


def dosyalari_yaz(hedef):
    for rel in sorted(PAYLOAD):
        yol = os.path.join(hedef, rel.replace('/', os.sep))
        try:
            n = yaz(yol, ''.join(PAYLOAD[rel]))
            log(True, 'file written: %s (%d bytes)' % (yol, n))
        except Exception as e:
            log(False, 'file %s: %s' % (yol, e))
    # write test: can we also delete/rename in that folder?
    try:
        t = os.path.join(hedef, '_yazma_testi.tmp')
        File.WriteAllText(t, 'x')
        File.Delete(t)
        log(True, 'folder is writable')
    except Exception as e:
        log(False, 'folder write test: %s' % e)


def excel_olustur(hedef):
    xl = None
    try:
        t = Type.GetTypeFromProgID('Excel.Application')
        if t is None:
            log(False, 'Excel is not installed (no Excel.Application COM object)')
            return
        xl = Activator.CreateInstance(t)
        log(True, 'Excel started via COM (version %s)' % xl.Version)
    except Exception as e:
        log(False, 'Excel COM start: %s' % e)
        return
    try:
        xl.DisplayAlerts = False
        xl.Visible = False
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets[1]
        ws.Name = 'Analysis Tracker'
        ws.Range('B2').Value2 = 'Analysis Tracker'
        ws.Range('B2').Font.Size = 20
        ws.Range('B2').Font.Bold = True
        ws.Range('B3').Value2 = 'Click a button. (created by analiz_takip_kurulum_test_16)'
        ws.Columns[1].ColumnWidth = 3
        ws.Columns[2].ColumnWidth = 40
        log(True, 'workbook and sheet created')
        yol = os.path.join(hedef, EXCEL_ADI)
        makro_var = False
        try:
            vbp = wb.VBProject
            kod = zlib.decompress(base64.b64decode(''.join(PAYLOAD['makro.txt']))).decode('utf-8-sig')
            mod = vbp.VBComponents.Add(1)
            mod.CodeModule.AddFromString(kod)
            makro_var = True
            log(True, 'VBA macro module inserted')
        except Exception as e:
            log(False, 'VBA macro insert (needs "Trust access to the VBA project object model"): %s' % e)
            bilgi('-> open makro.txt in the folder and paste it into Alt+F11 > Insert > Module by hand.')
        if makro_var:
            try:
                dugmeler = [('B5', 40, 'Open Analysis Tracker', 'ProgramiAc'),
                            ('B8', 26, 'Test: does a macro run?', 'TestMakro'),
                            ('B10', 26, 'Test: does PowerShell run?', 'TestPowerShell'),
                            ('B12', 26, 'Select Program Folder', 'KlasorDegistir')]
                for hucre, yuk, metin, makro in dugmeler:
                    r = ws.Range(hucre)
                    b = ws.Buttons().Add(r.Left, r.Top, r.Width, yuk)
                    b.OnAction = makro
                    b.Caption = metin
                log(True, '4 buttons created')
            except Exception as e:
                log(False, 'buttons: %s' % e)
        try:
            if File.Exists(yol):
                File.Delete(yol)
            wb.SaveAs(yol, 52 if makro_var else 51)     # 52 = xlsm, 51 = xlsx
            if makro_var:
                log(True, 'saved as .xlsm: %s' % yol)
            else:
                yol2 = yol.replace('.xlsm', '.xlsx')
                wb.SaveAs(yol2, 51)
                log(True, 'saved WITHOUT macro as .xlsx: %s' % yol2)
        except Exception as e:
            log(False, 'save: %s' % e)
        try:
            xl.Visible = True
            log(True, 'Excel window opened (left open for you)')
        except Exception as e:
            log(False, 'show Excel: %s' % e)
    except Exception as e:
        log(False, 'Excel step: %s' % e)
        bilgi(traceback.format_exc())
        try:
            xl.Quit()
        except Exception:
            pass


def rapor_goster(hedef):
    f = Form()
    f.Text = 'Analysis Tracker - setup test result'
    f.StartPosition = FormStartPosition.CenterScreen
    f.Size = Size(900, 520)
    t = TextBox()
    t.Multiline = True
    t.ReadOnly = True
    t.ScrollBars = ScrollBars.Both
    t.WordWrap = False
    t.Dock = DockStyle.Fill
    t.Font = Font('Consolas', 9.5)
    t.Text = '\r\n'.join(LOG)
    f.Controls.Add(t)
    b = Button()
    b.Text = 'Open folder'
    b.Dock = DockStyle.Bottom
    b.Height = 34

    def ac(s, e):
        try:
            System.Diagnostics.Process.Start('explorer.exe', '"' + hedef + '"')
        except Exception:
            pass
    b.Click += ac
    f.Controls.Add(b)
    f.Show()


def main():
    hedef = hedef_sec()
    if not hedef:
        return
    LOG.append('target folder: ' + hedef)
    LOG.append('time: ' + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    try:
        dosyalari_yaz(hedef)
        excel_olustur(hedef)
    except Exception as e:
        log(False, 'unexpected: %s' % e)
        bilgi(traceback.format_exc())
    try:
        File.WriteAllText(os.path.join(hedef, 'kurulum_test.log'), '\r\n'.join(LOG), UTF8Encoding(True))
    except Exception:
        pass
    rapor_goster(hedef)


import System
import System.Diagnostics
main()
