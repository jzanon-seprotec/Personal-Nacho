from bs4 import BeautifulSoup

def corregir_residues(input_file, output_file):
    with open(input_file, "r", encoding="utf-8") as f:
        html = f.read()

    soup = BeautifulSoup(html, "html.parser")

    for tr in soup.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) == 2 and "Residues" in tds[1].text:
            div = tds[1].find("div", class_="residues")
            if div and div.string:
                # Guardar texto
                texto = div.string.strip()
                # Eliminar celda incorrecta
                tds[1].decompose()
                # Crear nuevas celdas
                nueva_td_label = soup.new_tag("td")
                nueva_td_label.string = "Residues"

                nueva_td_contenido = soup.new_tag("td")
                nuevo_div = soup.new_tag("div", attrs={"class": "residues"})
                nuevo_div.string = texto
                nueva_td_contenido.append(nuevo_div)

                # Insertar después de la primera celda
                tds[0].insert_after(nueva_td_contenido)
                tds[0].insert_after(nueva_td_label)

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(str(soup))

    print(f"Archivo corregido guardado en: {output_file}")

# Ejemplo de uso
if __name__ == "__main__":
    archivo_entrada = "generated.html"               # Cambia si tienes otro nombre
    archivo_salida = "generated_column_fixed.html"   # Resultado
    corregir_residues(archivo_entrada, archivo_salida)
