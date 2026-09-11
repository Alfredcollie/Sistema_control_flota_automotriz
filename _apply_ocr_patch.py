# -*- coding: utf-8 -*-
# NOTA: script de migración YA APLICADO (queda como respaldo histórico).
# Sus plantillas ya usan el selector seguro compartido (dialogos_seguros.py),
# para no reintroducir el diálogo crudo de Tkinter que cierra la app en macOS.
path = r"C:\Users\Alberto\Desktop\Programa de control de flotilla automotriz para Win-Mac 210826\modulo_ventas.py"

with open(path, "r", encoding="utf-8") as f:
    content = f.read()

# 1) Import de Gemini (OCR con IA)
old_imp = 'try:\n    import pdfplumber\nexcept ImportError:\n    pdfplumber = None\n'
new_imp = ('try:\n    import pdfplumber\nexcept ImportError:\n    pdfplumber = None\n\n'
           'try:\n    from google import genai\n    from google.genai import types as genai_types\n'
           'except Exception:\n    genai = None\n    genai_types = None\n')
assert old_imp in content, "import anchor missing"
content = content.replace(old_imp, new_imp, 1)

# 2) Renombrar boton de autocompletado
content = content.replace('text="📄 Autocompletar desde PDF SUNAT"', 'text="📄 Cargar PDF y Autocompletar (OCR)"')

new_block = r'''    def autocompletar_desde_pdf(self):
        ruta = seleccionar_archivo_dialogo(
            titulo="Seleccionar Factura PDF",
            tipos=[("Archivos PDF", "*.pdf"), ("Imágenes", "*.png;*.jpg;*.jpeg")])
        if not ruta:
            return
        self._autocompletar_desde_archivo(ruta)

    def _autocompletar_desde_archivo(self, ruta):
        """Carga el PDF/imagen, extrae el texto o aplica OCR y rellena todos los campos."""
        try:
            self.bloquear_autocompletado_ruc = True
            rellenado = False
            texto = self._extraer_texto_pdf(ruta)
            if texto.strip():
                rellenado = self._rellenar_desde_texto(texto)
            if not rellenado:
                datos = self._ocr_factura_gemini(ruta)
                if datos:
                    rellenado = self._rellenar_desde_ocr(datos)

            self.ruta_archivo_temp = ruta
            if hasattr(self, "lbl_archivo"):
                self.lbl_archivo.configure(text=f"📎 {os.path.basename(ruta)}", text_color="#166534")
            self.actualizar_totales()
            self.al_seleccionar_cliente()

            if rellenado:
                messagebox.showinfo("Lectura de Factura", "Datos de la factura rellenados automáticamente (texto/OCR).")
            else:
                messagebox.showwarning("Lectura Incompleta", "No se pudieron leer automáticamente los datos del documento.\nRevise los campos manualmente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        finally:
            self.bloquear_autocompletado_ruc = False

    def _extraer_texto_pdf(self, ruta):
        """Extrae la capa de texto del PDF (pdfplumber y luego PyMuPDF). Devuelve '' si no hay texto."""
        texto = ""
        if pdfplumber is not None:
            try:
                with pdfplumber.open(ruta) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            texto += t + "\n"
            except Exception:
                pass
        if not texto.strip():
            try:
                import fitz
                doc = fitz.open(ruta)
                for page in doc:
                    texto += (page.get_text() or "") + "\n"
                doc.close()
            except Exception:
                pass
        return texto

    def _rellenar_desde_texto(self, texto):
        """Rellena los campos del formulario a partir del texto extraído. Devuelve True si rellenó algo."""
        relleno = False

        if re.search(r"FACTURA\s+ELECTR[OÓ]NICA", texto, re.IGNORECASE):
            self.combo_tipo.set("Factura (18% IGV)")
        elif re.search(r"BOLETA\s+DE\s+VENTA", texto, re.IGNORECASE):
            self.combo_tipo.set("Boleta (Sin IGV)")
        elif re.search(r"RECIBO\s+POR\s+HONORARIOS", texto, re.IGNORECASE):
            if re.search(r"Retenci[oó]n.*?IR[\s:\|]*\(?([\d\,\.]+)\)?", texto, re.IGNORECASE):
                self.combo_tipo.set("Recibo por Honorarios (8% Retención)")
            else:
                self.combo_tipo.set("Recibo por Honorarios (Sin Retención)")
        self.on_tipo_change(self.combo_tipo.get())

        nro_match = re.search(r"([EFB][0-9A-Z]{3}\s*-\s*\d+)", texto)
        if nro_match:
            self.ent_nro_doc.delete(0, tk.END); self.ent_nro_doc.insert(0, nro_match.group(1).replace(" ", ""))
            relleno = True

        fecha_match = re.search(r"Fecha de Emisi[oó]n\s*[:\-]?\s*(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto, re.IGNORECASE)
        if not fecha_match:
            fecha_match = re.search(r"(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto)
        if fecha_match:
            d, m, y = fecha_match.groups()
            fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
            self.ent_fecha.delete(0, tk.END)
            if fmt == "MM/DD/AAAA":
                self.ent_fecha.insert(0, f"{m}/{d}/{y}")
            else:
                self.ent_fecha.insert(0, f"{d}/{m}/{y}")
            relleno = True

        cliente_match = re.search(r"(?:Señor\(es\)|Señores|Razón Social|Cliente|Recibí\s*de)\s*[:\-]\s*(.+)", texto, re.IGNORECASE)
        rucs = re.findall(r"(?:RUC|R\.U\.C\.|Documento)\s*[:\-]?\s*(\d{11})", texto, re.IGNORECASE)

        if cliente_match:
            self.combo_cliente.set(cliente_match.group(1).strip())
            relleno = True
        if rucs:
            self.ent_ruc.configure(state="normal")
            self.ent_ruc.delete(0, tk.END)
            self.ent_ruc.insert(0, rucs[-1])
            relleno = True

        sub_m = re.search(r"(?:OP\.\s*GRAVADAS|SUB\s*TOTAL|Subtotal|Total por honorarios)[\s:S/\|]+([\d\,\.]+)", texto, re.IGNORECASE)
        tot_m = re.search(r"(?:IMPORTE\s*TOTAL|TOTAL\s*A\s*PAGAR|Total Neto Recibido)[\s:S/\|]+([\d\,\.]+)", texto, re.IGNORECASE)
        monto_base = 0.0
        if sub_m:
            try:
                monto_base = float(sub_m.group(1).replace(",", ""))
            except ValueError:
                monto_base = 0.0
        elif tot_m:
            try:
                t = float(tot_m.group(1).replace(",", ""))
                monto_base = t / 1.18 if "Factura" in self.combo_tipo.get() else t
            except ValueError:
                monto_base = 0.0
        if monto_base > 0:
            self.ent_subtotal.delete(0, tk.END); self.ent_subtotal.insert(0, f"{monto_base:.2f}")
            relleno = True

        if "Recibo" not in self.combo_tipo.get():
            det_match = re.search(r"(?:Detracci[oó]n|Porcentaje|Tasa).*?(\d{1,2}(?:\.\d{1,2})?)\s*%", texto, re.IGNORECASE)
            if not det_match:
                if re.search(r"Sujeta\s*a\s*detracci[oó]n", texto, re.IGNORECASE):
                    self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, CONFIG_REGIONAL.get("detraccion_porcentaje", "12"))
                    relleno = True
            else:
                self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, det_match.group(1))
                relleno = True

        return relleno

    def _ocr_factura_gemini(self, ruta):
        """Aplica OCR con IA (Google Gemini) a un PDF/imagen escaneado. Devuelve un dict con los campos o None."""
        if genai is None:
            return None
        imagenes = []
        try:
            import fitz
            doc = fitz.open(ruta)
            for page in doc:
                try:
                    pix = page.get_pixmap(dpi=200)
                    imagenes.append(pix.tobytes("png"))
                except Exception:
                    continue
            doc.close()
        except Exception:
            imagenes = []
        if not imagenes:
            try:
                with open(ruta, "rb") as f:
                    imagenes.append(f.read())
            except Exception:
                return None

        try:
            clave = os.environ.get("GEMINI_API_KEY", "").strip() or "AQ.Ab8RN6LTyHmVNUALwk6Wk7b2EMSzbZrVXVjg-cKUH7cSwnJ0Iw"
            cliente = genai.Client(api_key=clave)
        except Exception:
            return None

        prompt = (
            "Eres un lector OCR de facturas, boletas y recibos peruanos. "
            "Lee la imagen del comprobante y devuelve SOLO un JSON válido, sin texto adicional, con estas claves exactas:\n"
            '- "tipo": "factura", "boleta" o "recibo".\n'
            '- "numero_documento": serie y correlativo (ej. "F001-123"). Si no se ve, "".\n'
            '- "fecha": fecha en DD/MM/YYYY. Si no se ve, "".\n'
            '- "cliente": razón social o nombre del cliente. Si no se ve, "".\n'
            '- "ruc": exactamente 11 dígitos. Si no se ve o no son 11 dígitos, "".\n'
            '- "descripcion": concepto o descripción del servicio/venta. Si no se ve, "".\n'
            '- "subtotal": monto base sin IGV, número decimal (ej. 22212.00). Si no se ve, 0.\n'
            '- "igv": importe del IGV, número decimal. Si no se ve, 0.\n'
            '- "total": importe total, número decimal. Si no se ve, 0.\n'
            '- "detraccion": porcentaje de detracción o retención (ej. 12). Si no aplica, 0.\n'
            "Reglas: NO inventes datos. Montos como números con punto decimal, sin símbolo de moneda ni comas. "
            "Si un campo no se ve, devuélvelo vacío o 0."
        )

        try:
            partes = []
            if genai_types is not None:
                for img in imagenes[:3]:
                    partes.append(genai_types.Part.from_bytes(data=img, mime_type="image/png"))
            if not partes:
                return None
            respuesta = cliente.models.generate_content(model="gemini-3.5-flash-lite", contents=[prompt] + partes)
            texto_ia = (respuesta.text or "").strip()
        except Exception as e:
            print("⚠️ Error OCR Gemini:", e)
            return None

        m = re.search(r"\{.*\}", texto_ia, re.DOTALL)
        if not m:
            return None
        try:
            return json.loads(m.group(0))
        except Exception:
            return None

    def _rellenar_desde_ocr(self, datos):
        """Rellena los campos del formulario a partir del dict del OCR. Devuelve True si rellenó algo."""
        relleno = False
        try:
            tipo = str(datos.get("tipo") or "").lower()
            if "boleta" in tipo:
                self.combo_tipo.set("Boleta (Sin IGV)")
            elif "recibo" in tipo:
                self.combo_tipo.set("Recibo por Honorarios (8% Retención)")
            else:
                self.combo_tipo.set("Factura (18% IGV)")
            self.on_tipo_change(self.combo_tipo.get())

            nro = str(datos.get("numero_documento") or "").strip()
            if nro:
                self.ent_nro_doc.delete(0, tk.END); self.ent_nro_doc.insert(0, nro)
                relleno = True

            fecha = str(datos.get("fecha") or "").strip()
            mf = re.search(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", fecha)
            if mf:
                d, mes, anio = mf.groups()
                if len(anio) == 2:
                    anio = "20" + anio
                fmt = CONFIG_REGIONAL.get("formato_fecha", "DD/MM/AAAA")
                self.ent_fecha.delete(0, tk.END)
                if fmt == "MM/DD/AAAA":
                    self.ent_fecha.insert(0, f"{mes}/{d}/{anio}")
                else:
                    self.ent_fecha.insert(0, f"{d}/{mes}/{anio}")
                relleno = True

            cliente = str(datos.get("cliente") or "").strip()
            if cliente:
                self.combo_cliente.set(cliente)
                relleno = True

            ruc = re.sub(r"\D", "", str(datos.get("ruc") or ""))
            if len(ruc) == 11:
                self.ent_ruc.configure(state="normal")
                self.ent_ruc.delete(0, tk.END)
                self.ent_ruc.insert(0, ruc)
                relleno = True

            desc = str(datos.get("descripcion") or "").strip()
            if desc:
                self.ent_desc.delete(0, tk.END); self.ent_desc.insert(0, desc)
                relleno = True

            def _num(v):
                s = str(v).strip().replace("S/", "").replace("s/", "").replace("$", "").replace(" ", "")
                if "," in s and "." in s:
                    if s.rfind(",") > s.rfind("."):
                        s = s.replace(".", "").replace(",", ".")
                    else:
                        s = s.replace(",", "")
                elif "," in s:
                    s = s.replace(",", ".")
                try:
                    return float(s)
                except Exception:
                    return 0.0

            subtotal = _num(datos.get("subtotal"))
            total = _num(datos.get("total"))
            monto_base = subtotal
            if monto_base <= 0 and total > 0:
                monto_base = total / 1.18 if "Factura" in self.combo_tipo.get() else total
            if monto_base > 0:
                self.ent_subtotal.delete(0, tk.END); self.ent_subtotal.insert(0, f"{monto_base:.2f}")
                relleno = True

            det = _num(datos.get("detraccion"))
            if det > 0:
                self.ent_detraccion.delete(0, tk.END); self.ent_detraccion.insert(0, f"{det:g}")
                relleno = True
        except Exception as e:
            print("⚠️ Error rellenando desde OCR:", e)
        return relleno'''

new_adjuntar = r'''    def adjuntar_pdf_factura(self):
        """Adjunta manualmente el PDF de una factura local y autocompleta los campos (con OCR si es necesario)."""
        ruta = seleccionar_archivo_dialogo(
            titulo="Seleccionar PDF de la Factura",
            tipos=[("Archivos PDF", "*.pdf"),
                   ("Imágenes (JPG/PNG)", "*.png;*.jpg;*.jpeg"),
                   ("Todos los archivos", "*.*")])
        if not ruta:
            return
        self._autocompletar_desde_archivo(ruta)'''


def replace_method(content, name, block):
    lines = content.split("\n")
    start = None
    end = None
    for i, l in enumerate(lines):
        if l == "    def %s(self):" % name:
            start = i
        elif start is not None and l.startswith("    def ") and i > start:
            end = i
            break
    assert start is not None, "method %s not found" % name
    if end is None:
        end = len(lines)
    new_lines = block.split("\n")
    while new_lines and new_lines[-1] == "":
        new_lines.pop()
    return "\n".join(lines[:start] + new_lines + lines[end:])


content = replace_method(content, "autocompletar_desde_pdf", new_block)
content = replace_method(content, "adjuntar_pdf_factura", new_adjuntar)

with open(path, "w", encoding="utf-8", newline="\r\n") as f:
    f.write(content)
print("PATCH OK")
