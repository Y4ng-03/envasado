"""
admin_usuarios.py — Gestión segura de usuarios por consola (sin contraseñas en código).

Uso:
    python admin_usuarios.py listar
    python admin_usuarios.py crear <usuario> [rol: admin|supervisor|tv] [nombre completo]
    python admin_usuarios.py clave <usuario>
    python admin_usuarios.py estado <usuario> <activo|inactivo>
    python admin_usuarios.py eliminar <usuario>
"""

import getpass
import sys

from db_compat import ROLES_VALIDOS_USER, get_db, init_auth_tables, upsert_user


def main() -> int:
    if len(sys.argv) < 2 or sys.argv[1] not in ("listar", "crear", "clave", "estado", "eliminar"):
        print(__doc__)
        return 1

    cmd = sys.argv[1]
    with get_db() as conn:
        init_auth_tables(conn)

        if cmd == "listar":
            rows = conn.execute(
                "SELECT id, username, rol, nombre, activo, ultimo_login FROM usuarios ORDER BY id"
            ).fetchall()
            if not rows:
                print("No hay usuarios registrados todavía.")
                return 0
            for r in rows:
                print(
                    f"{r['id']:>3}  {r['username']:<16} rol={r['rol']:<10} activo={r['activo']}  {r['nombre']}  (últ. login: {r['ultimo_login'] or '-'})"
                )
            return 0

        if len(sys.argv) < 3:
            print("Falta el nombre de usuario.")
            return 1

        uname = sys.argv[2].strip().lower()

        if cmd == "crear":
            rol = (sys.argv[3] if len(sys.argv) >= 4 else "supervisor").strip().lower()
            if rol not in ROLES_VALIDOS_USER:
                print(f"Rol inválido '{rol}'. Usa: {', '.join(ROLES_VALIDOS_USER)}")
                return 1
            nombre = " ".join(sys.argv[4:]).strip() if len(sys.argv) >= 5 else uname
            p1 = getpass.getpass(f"Contraseña para '{uname}' (mín. 8 caracteres): ")
            p2 = getpass.getpass("Repite la contraseña: ")
            if p1 != p2:
                print("Las contraseñas no coinciden.")
                return 1
            if len(p1) < 8:
                print("Demasiado corta (mínimo 8 caracteres).")
                return 1
            upsert_user(conn, username=uname, password=p1, nombre=nombre, rol=rol)
            print(f"Usuario '{uname}' ({rol}) guardado e encriptado correctamente en la base de datos.")
            print(f"Para iniciar sesión en la web escribe en Usuario: {uname}")
            return 0

        if cmd == "eliminar":
            res = conn.execute("DELETE FROM usuarios WHERE LOWER(username) = ?", (uname,))
            if res.rowcount:
                print(f"Usuario '{uname}' eliminado de la base de datos.")
                return 0
            print(f"El usuario '{uname}' no existe.")
            return 1

        row = conn.execute(
            "SELECT id, username, nombre, rol, cargo, activo FROM usuarios WHERE LOWER(username) = ?",
            (uname,),
        ).fetchone()
        if not row:
            print(f"El usuario '{uname}' no existe.")
            return 1

        if cmd == "clave":
            p1 = getpass.getpass(f"Nueva contraseña para '{uname}' (mín. 8 caracteres): ")
            p2 = getpass.getpass("Repite la contraseña: ")
            if p1 != p2:
                print("Las contraseñas no coinciden.")
                return 1
            if len(p1) < 8:
                print("Demasiado corta (mínimo 8 caracteres).")
                return 1
            upsert_user(
                conn,
                username=uname,
                password=p1,
                nombre=row["nombre"],
                rol=row["rol"],
                cargo=row["cargo"],
                activo=row["activo"],
            )
            print(f"Contraseña de '{uname}' actualizada.")
            return 0

        if cmd == "estado":
            if len(sys.argv) < 4 or sys.argv[3].lower() not in ("activo", "inactivo"):
                print("Indica 'activo' o 'inactivo'.")
                return 1
            val = 1 if sys.argv[3].lower() == "activo" else 0
            conn.execute("UPDATE usuarios SET activo = ? WHERE id = ?", (val, row["id"]))
            print(f"Usuario '{uname}' marcado como {'activo' if val else 'inactivo'}.")
            return 0

    return 0


if __name__ == "__main__":
    sys.exit(main())
