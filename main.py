from modules.downloader import (
    descargar_licitaciones,
    filtrar_licitaciones,
    exportar_excel
)

from modules.scraper import abrir_licitacion
from modules.email_sender import enviar_correo

def main():

    MES = 4
    ANIO = 2025

    # DESCARGAR
    df = descargar_licitaciones(MES, ANIO)

    if df is None:
        return

    # FILTRAR
    df_filtrado = filtrar_licitaciones(
        df,
        monto_minimo=1000000
    )

    if df_filtrado is None:
        return

    # EXPORTAR TODO
    exportar_excel(
        df_filtrado,
        "licitaciones_filtradas.xlsx"
    )

    print("\nTOTAL FILTRADAS:")
    print(len(df_filtrado))

    print("\nTOTAL FILTRADAS:")
    print(len(df_filtrado))

    print("\nINICIANDO SCRAPING MASIVO...")

    for i in range(len(df_filtrado)):

        try:

            fila = df_filtrado.iloc[i]

            codigo_real = fila["CodigoExterno"]

            empresa = fila["NombreProveedor"]

            rut = fila["RutProveedor"]

            organismo = fila["NombreOrganismo"]

            print("\n======================")
            print(f"LICITACION {i + 1}")
            print("======================")

            print("CODIGO:")
            print(codigo_real)

            print("EMPRESA:")
            print(empresa)

            print("RUT:")
            print(rut)

            abrir_licitacion(
                codigo_real,
                empresa,
                rut,
                organismo
            )

        except Exception as e:

            print("\nERROR EN LICITACION:")
            print(e)


if __name__ == "__main__":
    main()