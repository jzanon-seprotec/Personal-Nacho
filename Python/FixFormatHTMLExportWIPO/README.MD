This macro fixes the format of generated.html from WIPO.

Use: 
path of source and target.

# Ejemplo de uso
if __name__ == "__main__":
    archivo_entrada = "D:\Borrar_37_XMLyHTML\Borrar\generated.html"               # Cambia si tienes otro nombre
    archivo_salida = "D:\Borrar_37_XMLyHTML\Borrar\generated_column_fixed.html"   # Resultado
    corregir_residues(archivo_entrada, archivo_salida)
