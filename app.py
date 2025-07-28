from flask import Flask, request, jsonify, send_file
import os
import zipfile
import tempfile
from auditor import ejecutar_auditoria

app = Flask(__name__)

@app.route("/auditar", methods=["POST"])
def auditar():
    try:
        # Crear carpeta temporal para trabajar
        temp_dir = tempfile.mkdtemp()
        info_auditada_path = os.path.join(temp_dir, "Información Auditada")
        fv_folder = os.path.join(info_auditada_path, "FVs Auditados")
        os.makedirs(fv_folder)

        # Guardar archivos BOM y CRUCE
        bom_file = request.files.get("bom")
        cruce_file = request.files.get("cruce")
        zip_pdfs = request.files.get("pdfs")

        if not bom_file or not cruce_file or not zip_pdfs:
            return jsonify({"error": "Faltan archivos: se requieren 'bom', 'cruce' y 'pdfs' (ZIP)"}), 400

        bom_path = os.path.join(info_auditada_path, bom_file.filename)
        cruce_path = os.path.join(info_auditada_path, cruce_file.filename)
        zip_path = os.path.join(temp_dir, "fvs.zip")

        bom_file.save(bom_path)
        cruce_file.save(cruce_path)
        zip_pdfs.save(zip_path)

        # Descomprimir los PDFs dentro de la carpeta FVs Auditados
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(fv_folder)

        # Ejecutar auditoría
        out1, out2 = ejecutar_auditoria(temp_dir)

        # Comprimir resultados en un ZIP para devolver
        output_zip = os.path.join(temp_dir, "resultado_auditoria.zip")
        with zipfile.ZipFile(output_zip, 'w') as zipf:
            zipf.write(out1, arcname=os.path.basename(out1))
            zipf.write(out2, arcname=os.path.basename(out2))

        return send_file(output_zip, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
