import json
import os

import pytest

from infraestructura import administradoras


def test_leer_formato_con_clave(ruta_administradoras):
    raw, clave, registros, mtime = administradoras.leer(ruta_administradoras)
    assert clave == "administradoras" and len(registros) == 7 and mtime > 0


def test_leer_lista_directa(tmp_path):
    ruta = tmp_path / "a.json"
    ruta.write_text(json.dumps([{"NOMBRE": "X", "NIT": "1"}]), encoding="utf-8")
    assert administradoras.leer(str(ruta))[1:3] == (None, [{"NOMBRE": "X", "NIT": "1"}])


def test_leer_formato_invalido(tmp_path):
    ruta = tmp_path / "a.json"
    ruta.write_text(json.dumps({"x": 1}), encoding="utf-8")
    with pytest.raises(ValueError):
        administradoras.leer(str(ruta))


def test_guardar_con_respaldo(ruta_administradoras):
    raw, clave, registros, mtime = administradoras.leer(ruta_administradoras)
    registros[0]["CUENTA"] = "237005099"
    nuevo = administradoras.guardar(ruta_administradoras, raw, clave, registros)
    assert os.path.exists(ruta_administradoras + ".bak")
    assert administradoras.leer(ruta_administradoras)[2][0]["CUENTA"] == "237005099"
    assert not administradoras.modificado_desde(ruta_administradoras, nuevo)


def test_tabla_normalizada(ruta_administradoras):
    tabla = administradoras.tabla(ruta_administradoras)
    assert list(tabla.columns) == ["nombre", "nit", "categoria", "cuenta", "codigo"]
    assert tabla.loc[tabla["codigo"] == "14-11", "cuenta"].iloc[0] == "237006001"