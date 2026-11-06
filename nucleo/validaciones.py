import re

PATRON_COMPROBANTE = re.compile(r"^[A-Z0-9]{1,5}$")
PATRON_CUENTA = re.compile(r"^\d{4,12}$")
PATRON_TERCERO = re.compile(r"^\d{5,15}(?:-\d)?$")
PATRON_TERCERO_REFERENCIA = re.compile(r"^\d{6,15}(?:-\d)?$")
PATRON_NIT = re.compile(r"^\d{5,15}$")
PATRON_CODIGO_PILA = re.compile(r"^[A-Z0-9-]{2,10}$")

MSG_COMPROBANTE = "El comprobante debe tener de 1 a 5 caracteres alfanuméricos (ej. EG1)."
MSG_CUENTA = "La cuenta contable debe tener de 4 a 12 dígitos."
MSG_TERCERO = "El tercero debe ser un NIT o cédula numérico (ej. 900123456 o 900123456-7)."
MSG_NIT = "El NIT debe tener de 5 a 15 dígitos, sin dígito de verificación."
MSG_CODIGO_PILA = "El código PILA admite letras, números y guion (ej. EPS010, 25-14)."