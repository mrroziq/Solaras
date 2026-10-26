from flask import Flask, request, jsonify
from pathlib import Path
import pandas as pd
import json

app = Flask(__name__)


@app.route("/api/upload", methods=["POST"])
def upload_dataset():
    try:
        if "file" not in request.files:
            return jsonify({
                "error": "File tidak ditemukan."
            }), 400

        file = request.files["file"]

        if not file.filename:
            return jsonify({
                "error": "Pilih file terlebih dahulu."
            }), 400

        extension = Path(file.filename).suffix.lower()

        if extension == ".csv":
            df = pd.read_csv(file)
        elif extension in [".xlsx", ".xls"]:
            df = pd.read_excel(file)
        else:
            return jsonify({
                "error": "Format file tidak didukung."
            }), 400

        df.columns = [str(col).strip() for col in df.columns]

        if df.empty or len(df.columns) < 2:
            return jsonify({
                "error": "Dataset harus memiliki minimal dua kolom dan satu baris data."
            }), 400

        if any(not col for col in df.columns):
            return jsonify({
                "error": "Nama kolom tidak boleh kosong."
            }), 400

        if not df.columns.is_unique:
            return jsonify({
                "error": "Nama kolom tidak boleh duplikat."
            }), 400

        object_column = df.columns[0]
        parameters = []

        for col in df.columns[1:]:
            series = df[col]
            nonempty = series.dropna()

            if pd.api.types.is_numeric_dtype(series):
                parameter = {
                    "name": col,
                    "type": "numeric",
                    "min": float(nonempty.min()) if not nonempty.empty else 0,
                    "max": float(nonempty.max()) if not nonempty.empty else 0
                }
            else:
                parameter = {
                    "name": col,
                    "type": "categorical",
                    "values": [
                        str(value)
                        for value in nonempty.unique()
                    ]
                }

            parameters.append(parameter)

        data = json.loads(
            df.to_json(
                orient="records",
                force_ascii=False
            )
        )

        objects = [
            str(value)
            for value in df[object_column].tolist()
        ]

        return jsonify({
            "success": True,
            "rows": len(df),
            "columns": len(parameters),
            "missing": int(df.isna().sum().sum()),
            "object_column": object_column,
            "objects": objects,
            "parameters": parameters,
            "data": data
        })

    except Exception as error:
        app.logger.exception("Gagal memproses dataset")
        return jsonify({
            "error": f"Gagal memproses dataset: {str(error)}"
        }), 500


if __name__ == "__main__":
    app.run(debug=True)
