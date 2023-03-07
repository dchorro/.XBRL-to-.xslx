import argparse
import re
import traceback
import os
import xlsxwriter
import zipfile


# Probando nueva regex para ver si extrae todos los datos.
estados_financieros_rgx = re.compile(r"""<pgc-[\w\d-]*:(\w+) [\w\d="]+ contextRef="(\w.\w+)" unitRef="\w+"[\w\d\_\=\s\"]*>(-?\d*\.{0,1}\d+)</pgc-[\w\d-]*:[\w]+>""")


pgc_regex       = re.compile(r"""<link:schemaRef xlink:type="simple" xlink:href="http://www.icac.meh.es/taxonomia/[\w\d-]+/pgc07-([\w]+[-]{0,1}[\w]+).xsd" />""")
instant_regex   = re.compile(r"<xbrli:instant>([\w-]+)</xbrli:instant>")
dates_regex     = re.compile(r"<xbrli:startDate>([\w-]+)</xbrli:startDate>|<xbrli:endDate>([\w-]+)</xbrli:endDate>")

nif_regex       = re.compile(r"""<dgi-est-gen:IdentifierValue contextRef="D.ACTUAL">([\d\w]+)</dgi-est-gen:IdentifierValue>""")

# name_regex_og      = re.compile(r"""<dgi-est-gen:LegalNameValue contextRef="D.ACTUAL">([\d\w\.\-\,\&\;\(\) ]+)</dgi-est-gen:LegalNameValue>""")
# name_regex      = re.compile(r"""<dgi-est-gen:LegalNameValue contextRef="D.ACTUAL">([\d\w\.\-\,\& ]+)</dgi-est-gen:LegalNameValue>""")
name_regex      = re.compile(r"""<dgi-est-gen:LegalNameValue contextRef="D.ACTUAL">([^<>]+)<\/dgi-est-gen:LegalNameValue>""")

# city_regex      = re.compile(r"""<dgi-est-gen:MunicipalityName contextRef="D.ACTUAL">([\w\'\/ \-()]+)</dgi-est-gen:MunicipalityName>""")
city_regex      = re.compile(r"""<dgi-est-gen:MunicipalityName contextRef="D.ACTUAL"[^<>]*>([^<>]+)</dgi-est-gen:MunicipalityName>""")

cp_regex        = re.compile(r"""<dgi-est-gen:ZipPostalCode contextRef="D.ACTUAL">([\d]+)</dgi-est-gen:ZipPostalCode>""")



taxonomies = ['normal', 'pymes', 'abreviado-completo', 'abreviado', 'mixto', 'pymes-completo']



fields = {
    "ActivoNoCorriente" : 0,
    "ActivoNoCorrienteInmovilizadoIntangible" : 1,
    "ActivoNoCorrienteInmovilizadoMaterial" : 2,
    "ActivoNoCorrienteInversionesInmobiliarias" : 3,
    "ActivoNoCorrienteInversionesEmpresasGrupoEmpresasAsociadasLargoPlazo" : 4,
    "ActivoNoCorrienteInversionesFinancierasLargoPlazo" : 5,
    "ActivoNoCorrienteActivosImpuestoDiferido" : 6,
    "ActivoNoCorrienteDeudasComercialesNoCorriente" : 7,
    "ActivoCorriente" : 8,
    "ActivoCorrienteExistencias" : 9,
    "ActivoCorrienteDeudoresComercialesOtrasCuentasCobrar" : 10,
    "ActivoCorrienteDeudoresComercialesOtrasCuentasCobrarClientesVentasPrestacionesServicios" : 11,
    "ActivoCorrienteDeudoresComercialesOtrasCuentasCobrarClientesVentasPrestacionesServiciosLargoPlazo" : 12,
    "ActivoCorrienteDeudoresComercialesOtrasCuentasCobrarClientesVentasPrestacionesServiciosCortoPlazo" : 13,
    "ActivoCorrienteDeudoresComercialesOtrasCuentasCobrarAccionistasDesembolsosExigidos" : 14,
    "ActivoCorrienteDeudoresComercialesOtrasCuentasCobrarOtrosDeudores" : 15,
    "ActivoCorrienteInversionesEmpresasGrupoEmpresasAsociadasCortoPlazo" : 16,
    "ActivoCorrienteInversionesFinancierasCortoPlazo" : 17,
    "ActivoCorrientePeriodificacionesCortoPlazo" : 18,
    "ActivoCorrienteEfectivoOtrosActivosLiquidosEquivalentes" : 19,
    "TotalActivo" : 20,
    "PatrimonioNeto" : 21,
    "PatrimonioNetoFondosPropios" : 22,
    "PatrimonioNetoFondosPropiosCapital" : 23,
    "PatrimonioNetoFondosPropiosCapitalEscriturado" : 24,
    "PatrimonioNetoFondosPropiosCapitalNoExigido" : 25,
    "PatrimonioNetoFondosPropiosPrimaEmision" : 26,
    "PatrimonioNetoFondosPropiosReservas" : 27,
    "PatrimonioNetoFondosPropiosReservasReservaCapitalizacion" : 28,
    "PatrimonioNetoFondosPropiosReservasOtrasReservas" : 29,
    "PatrimonioNetoFondosPropiosAccionesParticipacionesPatrimonioPropias" : 30,
    "PatrimonioNetoFondosPropiosResultadosEjerciciosAnteriores" : 31,
    "PatrimonioNetoFondosPropiosOtrasAportacionesSocios" : 32,
    "PatrimonioNetoFondosPropiosResultadoEjercicio" : 33,
    "PatrimonioNetoFondosPropiosDividendoCuenta" : 34,
    "PatrimonioNetoAjustesCambioValor" : 35,
    "PatrimonioNetoSubvencionesDonacionesLegadosRecibidos" : 36,
    "PasivoNoCorriente" : 37,
    "PasivoNoCorrienteProvisionesLargoPlazo" : 38,
    "PasivoNoCorrienteDeudasLargoPlazo" : 39,
    "PasivoNoCorrienteDeudasLargoPlazoDeudasEntidadesCredito" : 40,
    "PasivoNoCorrienteDeudasLargoPlazoAcreedoresArrendamientoFinanciero" : 41,
    "PasivoNoCorrienteDeudasLargoPlazoOtrasDeudas" : 42,
    "PasivoNoCorrienteDeudasEmpresasGrupoEmpresasAsociadasLargoPlazo" : 43,
    "PasivoNoCorrientePasivosImpuestoDiferido" : 44,
    "PasivoNoCorrientePeriodificacionesLargoPlazo" : 45,
    "PasivoNoCorrienteAcreedoresComercialesNoCorrientes" : 46,
    "PasivoNoCorrienteDeudaCaracteristicasEspecialesLargoPlazo" : 47,
    "PasivoCorriente" : 48,
    "PasivoCorrienteProvisionesCortoPlazo" : 49,
    "PasivoCorrienteDeudasCortoPlazo" : 50,
    "PasivoCorrienteDeudasCortoPlazoDeudasEntidadesCredito" : 51,
    "PasivoCorrienteDeudasCortoPlazoAcreedoresArrendamientoFinanciero" : 52,
    "PasivoCorrienteDeudasCortoPlazoOtrasDeudas" : 53,
    "PasivoCorrienteDeudasEmpresasGrupoEmpresasAsociadas" : 54,
    "PasivoCorrienteAcreedoresComercialesOtrasCuentasPagar" : 55,
    "PasivoCorrienteAcreedoresComercialesOtrasCuentasPagarProveedores" : 56,
    "PasivoCorrienteAcreedoresComercialesOtrasCuentasPagarProveedoresLargoPlazo" : 57,
    "PasivoCorrienteAcreedoresComercialesOtrasCuentasPagarProveedoresCortoPlazo" : 58,
    "PasivoCorrienteAcreedoresComercialesOtrasCuentasPagarOtrosAcreedores" : 59,
    "PasivoCorrientePeriodificacionesCortoPlazo" : 60,
    "PasivoCorrienteDeudasCaracteristicasEspecialesCortoPlazo" : 61,
    "PatrimonioNetoPasivoTotal" : 62,
    "PerdidasGananciasOperacionesContinuadasImporteNetoCifraNegocios" : 63,
    "PerdidasGananciasOperacionesContinuadasVariacionExistenciasProductosTerminadosProductosCursoFabricacion" : 64,
    "PerdidasGananciasOperacionesContinuadasTrabajosRealizadosEmpresaActivo" : 65,
    "PerdidasGananciasOperacionesContinuadasAprovisionamientos" : 66,
    "PerdidasGananciasOperacionesContinuadasOtrosIngresosExplotacion" : 67,
    "PerdidasGananciasOperacionesContinuadasGestionPersonal" : 68,
    "PerdidasGananciasOperacionesContinuadasOtrosGastosExplotacion" : 69,
    "PerdidasGananciasOperacionesContinuadasAmortizacionInmovilizado" : 70,
    "PerdidasGananciasOperacionesContinuadasImputacionSubvencionesInmovilizadoNoFinancieroOtras" : 71,
    "PerdidasGananciasOperacionesContinuadasExcesosProvisiones" : 72,
    "PerdidasGananciasOperacionesContinuadasDeterioroResultadoEnajenacionesInmovilizado" : 73,
    "PerdidasGananciasOtrosResultados" : 74,
    "PerdidasGananciasResultadoExplotacion" : 75,
    "PerdidasGananciasOperacionesContinuadasIngresosFinancieros" : 76,
    "PerdidasGananciasOperacionesContinuadasIngresosFinancierosImputacionSubvencionesDonacionesLegadosCaracterFinanciero" : 77,
    "PerdidasGananciasOperacionesContinuadasIngresosFinancierosOtrosIngresosFinancieros" : 78,
    "PerdidasGananciasOperacionesContinuadasGastosFinancieros" : 79,
    "PerdidasGananciasOperacionesContinuadasVariacionValorRazonableInstrumentosFinancieros" : 80,
    "PerdidasGananciasOperacionesContinuadasDiferenciasCambio" : 81,
    "PerdidasGananciasOperacionesContinuadasDeterioroResultadoEnajenacionesInstrumentosFinancieros" : 82,
    "PerdidasGananciasOperacionesContinuadasOtrosIngresosGastosCaracterFinanciero" : 83,
    "PerdidasGananciasOperacionesContinuadasOtrosIngresosGastosCaracterFinancieroIncorporacionActivoGastosFinancieros" : 84,
    "PerdidasGananciasOperacionesContinuadasOtrosIngresosGastosCaracterFinancieroIngresosFinancierosDerivadosConveniosAcreedores" : 85,
    "PerdidasGananciasOperacionesContinuadasOtrosIngresosGastosCaracterFinancieroRestoIngresosGastos" : 86,
    "PerdidasGananciasResultadoFinanciero" : 87,
    "PerdidasGananciasResultadoAntesImpuestos" : 88,
    "PerdidasGananciasOperacionesContinuadasImpuestosSobreBeneficios" : 89,
    "PerdidasGananciasResultadoEjercicio" : 90
}


def get_taxonomy(filename):
    with open(filename+"\\DEPOSITO.xbrl", encoding="utf8") as f:
        file = f.read()    

    pgc = re.findall(pgc_regex, file)[0]

    return pgc


def final_func(paths, output_name):
    if len(paths) == 0:
        raise Exception("No se han detectado documentos.")
    
    off_row, off_col = 0, 5

    workbook = xlsxwriter.Workbook(output_name + '.xlsx')

    worksheet = workbook.add_worksheet("Cuentas_Cooperativas")
    

    cell_format = workbook.add_format()
    cell_format.set_align('center')
    
    # write(row, column)
    worksheet.write(1, 0, "Cooperativa", cell_format)
    worksheet.write(1, 1, "Ciudad", cell_format)
    worksheet.write(1, 2, "CIF", cell_format)
    worksheet.write(1, 3, "Código postal", cell_format)
    worksheet.write(1, 4, "Año", cell_format)
    

    for key in fields:
        worksheet.write(1, off_col + fields[key], key, cell_format)

    written = {}

    for filename in paths:
        with open(filename+"\\DEPOSITO.xbrl", encoding="utf8") as f:
            file = f.read()

        try:
        
            instant_list = re.findall(instant_regex, file)

            if len(instant_list) >= 2:
                i_actual, i_anterior = instant_list[0], instant_list[1]
            else:
                raise Exception("Ha habido un error, en el formato del documento debe dehaber al menos 2 instantes de fechas para I.ACTUAL, I.ANTERIOR")

            dates_list = re.findall(dates_regex, file)

            dates_list = [''.join(x).split()[0] for x in dates_list]

            if len(dates_list) >= 4:
                start_date_actual, end_date_actual = dates_list[0], dates_list[1]
                start_date_anterior, end_date_anterior = dates_list[2], dates_list[3]
            else:
                raise Exception("Ha habido un error, en el formato del documento debe dehaber al menos 4 fechas para D.ACTUAL start y end, D.ANTERIOR start y end")

            estados_financieros_list = re.findall(estados_financieros_rgx, file)

            nif = re.findall(nif_regex, file)[0]
            name = re.findall(name_regex, file)[0]
            city = re.findall(city_regex, file)[0]
            city = city.upper()
            cp_list = re.findall(cp_regex, file)
            if len(cp_list) > 0:
                cp = cp_list[0]
            anyo_actual = i_actual.split("-")[0]
            anyo_anterior = i_anterior.split("-")[0]

            
            # Ya hemos escrito al menos una vez esa empresa, entonces escribir solo anterior
            if nif in written:
                worksheet.write(2 + off_row, 0, name, cell_format)

                worksheet.write(2 + off_row, 1, city, cell_format)

                worksheet.write(2 + off_row, 2, nif, cell_format)

                worksheet.write(2 + off_row, 3, cp, cell_format)
                
                worksheet.write(2 + off_row, 4, int(anyo_anterior), cell_format)

                for key in fields:
                    worksheet.write(2 + off_row, off_col + fields[key], 0, cell_format)
            
            
                for elem in estados_financieros_list:
                    apartado, contexto, cantidad = elem

                    try:
                        fields[apartado]
                    except:
                        continue
                    if contexto == "I.ACTUAL" or contexto == "D.ACTUAL":
                        pass
                    else:
                        worksheet.write(2 + off_row, off_col + fields[apartado], float(cantidad), cell_format)

                off_row += 1
                
            else:
                written[nif] = 0
                
                worksheet.write(2 + off_row, 0, name, cell_format)
                worksheet.write(3 + off_row, 0, name, cell_format)

                worksheet.write(2 + off_row, 1, city, cell_format)
                worksheet.write(3 + off_row, 1, city, cell_format)

                worksheet.write(2 + off_row, 2, nif, cell_format)
                worksheet.write(3 + off_row, 2, nif, cell_format)

                worksheet.write(2 + off_row, 3, cp, cell_format)
                worksheet.write(3 + off_row, 3, cp, cell_format)
                
                worksheet.write(2 + off_row, 4, int(anyo_actual), cell_format)
                worksheet.write(3 + off_row, 4, int(anyo_anterior), cell_format)

                for key in fields:                    
                    worksheet.write(2 + off_row, off_col + fields[key], 0, cell_format)
                    worksheet.write(3 + off_row, off_col + fields[key], 0, cell_format)
            
            
                for elem in estados_financieros_list:
                    apartado, contexto, cantidad = elem

                    try:
                        fields[apartado]
                    except:
                        continue
                    if contexto == "I.ACTUAL" or contexto == "D.ACTUAL":
                        worksheet.write(2 + off_row, off_col + fields[apartado], float(cantidad), cell_format)
                    else:
                        worksheet.write(3 + off_row, off_col + fields[apartado], float(cantidad), cell_format)

                off_row += 2

        except Exception as e:
            traceback.print_exc()
            print(filename)
            break            

    worksheet.autofit()
    workbook.close()
    f.close()
    return



def parse_file(filename):
    with open(filename+"\\DEPOSITO.xbrl", encoding="utf8") as f:
        file = f.read()

    # pgc = re.findall(pgc_regex, file)[0]

    instant_list = re.findall(instant_regex, file)

    if len(instant_list) >= 2:
        i_actual, i_anterior = instant_list[0], instant_list[1]
    else:
        raise Exception("Ha habido un error, en el formato del documento debe dehaber al menos 2 instantes de fechas para I.ACTUAL, I.ANTERIOR")


    dates_list = re.findall(dates_regex, file)

    dates_list = [''.join(x).split()[0] for x in dates_list]

    if len(dates_list) >= 4:
        start_date_actual, end_date_actual = dates_list[0], dates_list[1]
        start_date_anterior, end_date_anterior = dates_list[2], dates_list[3]
    else:
        raise Exception("Ha habido un error, en el formato del documento debe dehaber al menos 4 fechas para D.ACTUAL start y end, D.ANTERIOR start y end")

    # print(start_date_actual, end_date_actual)
    # print(start_date_anterior, end_date_anterior)


    estados_financieros_list = re.findall(estados_financieros_rgx, file)
    
    
    nif = re.findall(nif_regex, file)[0]
    name = re.findall(name_regex, file)[0]
    city = re.findall(city_regex, file)[0]
    cp_list = re.findall(cp_regex, file)
    if len(cp_list) > 0:
        cp = cp_list[0]
    anyo_actual = i_actual.split("-")[0]
    anyo_anterior = i_anterior.split("-")[0]

    # print(f"Año actual {anyo_actual}    Año anterior {anyo_anterior}")

    # if name == "SAT EL CUARTAZO" and (anyo_actual == "2020"):
    #     print(filename)
    #     return
    # else:
    #     return

    row, col = 0, 0

    off_col = 5

    # workbook = xlsxwriter.Workbook('Cuentas_' + str(nif) + "_"  + str(anyo_actual) + "_" + str(anyo_anterior) + '.xlsx')

    # workbook = xlsxwriter.Workbook(filename+'Cuentas_' + str(nif) + "_"  + str(anyo_actual) + "_" + str(anyo_anterior) + '.xlsx')
    
    # workbook = xlsxwriter.Workbook('Cuentas_' + str(nif) + "_"  + str(anyo_anterior) + '.xlsx')

    # worksheet = workbook.add_worksheet("Test_Balance")
    
    #   worksheet.write(row, column, content)

    # worksheet.write(1, 0, "Cooperativa")
    # worksheet.write(2, 0, name)
    # worksheet.write(3, 0, name)

    # worksheet.write(1, 1, "Ciudad")
    # worksheet.write(2, 1, city)
    # worksheet.write(3, 1, city)

    # worksheet.write(1, 2, "CIF")
    # worksheet.write(2, 2, nif)
    # worksheet.write(3, 2, nif)

    # worksheet.write(1, 3, "Código postal")
    # worksheet.write(2, 3, cp)
    # worksheet.write(3, 3, cp)

    # worksheet.write(1, 4, "Año")
    # worksheet.write(2, 4, anyo_actual)
    # worksheet.write(3, 4, anyo_actual)

    # for key in fields:
    #     worksheet.write(1, off_col + fields[key], key)
    
    print(f"Nombre: {name}, Año actual: {anyo_actual}, Año anterior: {anyo_anterior}")  
    
    for elem in estados_financieros_list:
        apartado, contexto, cantidad = elem

        # fields es el diccionario con todos los campos del archivo
        
        # Campos extra: Nombre de la cooperativa, Identificador de la cooperativa, código postal, 
        # ciudad, año de las cuentas

        # print(apartado, contexto, cantidad)

        if apartado == "PerdidasGananciasOperacionesContinuadasImporteNetoCifraNegocios":
            print(apartado, contexto, cantidad)

        # try:
        #     # worksheet.write(1, off_col + fields[apartado], apartado)
        #     fields[apartado]
        # except:
        #     continue
        # if contexto == "I.ACTUAL":
            # print("worksheet.write(1, " + str(fields[apartado]) + ", " + str(cantidad) + ")")
            # worksheet.write(2, off_col + fields[apartado], float(cantidad))
        # else:
            # print("worksheet.write(2, " + str(fields[apartado]) + ", " + str(cantidad) + ")")
            # worksheet.write(3, off_col + fields[apartado], float(cantidad))
        # print()
        
    
    
    # workbook.close()
    f.close()
    
    pass


def path_to_files(origin):
    res = []
    def fast_scandir(dirname):
        for f in os.scandir(dirname):
            if f.is_dir():
                fast_scandir(f.path)
            elif f.name == "DEPOSITO.xbrl" or f.name == "deposito.xbrl":
                res.append(f.path[:-13])
    fast_scandir(origin)
    return res


def list_files_taxonomy(paths):
    path_taxonomies_normal, path_taxonomies_pymes, path_taxonomies_abreviado_completo, path_taxonomies_abreviado, path_taxonomies_mixto, path_taxonomies_pymes_completo = [], [], [], [], [], []
    
    for filename in paths:
        with open(filename+"\\DEPOSITO.xbrl", encoding="utf8") as f:
            file = f.read()
    
        pgc_regex = re.compile(r"""<link:schemaRef xlink:type="simple" xlink:href="http://www.icac.meh.es/taxonomia/[\w\d-]+/pgc07-([\w]+[-]{0,1}[\w]+).xsd" />""")

        pgc = re.findall(pgc_regex, file)[0]

        if pgc == 'normal':
            path_taxonomies_normal.append(filename)
        elif pgc == 'pymes':
            path_taxonomies_pymes.append(filename)
        elif pgc == 'abreviado-completo':
            path_taxonomies_abreviado_completo.append(filename)
        elif pgc == 'abreviado':
            path_taxonomies_abreviado.append(filename)
        elif pgc == 'mixto':
            path_taxonomies_mixto.append(filename)
        else:
            path_taxonomies_pymes_completo.append(filename)
            

    return path_taxonomies_normal, path_taxonomies_pymes, path_taxonomies_abreviado_completo, path_taxonomies_abreviado, path_taxonomies_mixto, path_taxonomies_pymes_completo


def unzip_folders(origin):
    res = []
    def fast_scandir(dirname):
        for f in os.scandir(dirname):
            if f.is_dir():
                fast_scandir(f.path)
            elif f.name.count(".zip") > 0 or f.name.count(".ZIP") > 0:
                res.append(f.path)
            # elif f.name.count(".pdf") > 0:
                # os.remove(f.path)
                # pass
    fast_scandir(origin)
    
    for p in res:
        with zipfile.ZipFile(p, 'r') as zip_ref:
            zip_ref.extractall(p[:-4])
    


if __name__ == '__main__':
    # taxonomies = ['normal', 'pymes', 'abreviado-completo', 'abreviado', 'mixto', 'pymes-completo']
    
    # unzip_folders("datos\\RM Castellón documentos")

    # PATHS VALENCIA
    # paths = path_to_files("datos\\Coop_agroalimentarias")
    # paths.reverse()
    # final_func(paths, "Cuentas_Valencia")

    
    # PATHS CASTELLON
    # paths = path_to_files("datos\\RM Castellón")
    # paths.reverse()
    # final_func(paths, "Cuentas_Castellon")
    
    # PATHS ALICANTE
    # paths = path_to_files("datos\\RM_Alicante")
    # paths.reverse()
    # final_func(paths, "Cuentas_Alicante")


    parser = argparse.ArgumentParser(description='Transformar cuentas de formato .XBRL a .xlsx')
    parser.add_argument('dircuentas', metavar='dircuentas', type=str,
                        help='Carpeta con las cuentas de la cooperativa')

    parser.add_argument('nombre', metavar='nombre', type=str,
                        help='Nombre del archivo .xlsx donde se escribiran las cuentas.')

    parser.add_argument('-z', '--zip', dest='zip', action='store_true', default=False, 
                    help='Flag que activa la opción de descomprimir los archivos .zip en la carpeta dircuentas.')

    args = parser.parse_args()

    origin = args.dircuentas
    dest_name = args.nombre

    if args.zip:
        unzip_folders(origin)

    paths = path_to_files(origin)
    paths.reverse()

    final_func(paths, dest_name)