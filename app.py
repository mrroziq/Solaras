from flask import Flask, render_template, request, jsonify, send_from_directory
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

app = Flask(
    __name__,
    template_folder=str(BASE_DIR),
    static_folder=str(BASE_DIR),
    static_url_path="/"
)

def build_soft_set(df, rules, object_col):
    soft_set = {}

    for parameter, rule in rules.items():
        if rule["type"] == "numeric":
            value = float(rule["value"])
            operator = rule["operator"]

            if operator == "≥":
                mask = df[parameter] >= value
            elif operator == "≤":
                mask = df[parameter] <= value
            elif operator == ">":
                mask = df[parameter] > value
            elif operator == "<":
                mask = df[parameter] < value
            elif operator == "==":
                mask = df[parameter] == value
            else:
                mask = pd.Series(False, index=df.index)
        else:
            values = rule.get("values", [])
            mask = df[parameter].astype(str).isin(values)

        members = df.loc[mask, object_col].dropna().astype(str).tolist()
        soft_set[parameter] = members

    return soft_set


@app.route("/assets/<path:filename>")
def assets(filename):
    return send_from_directory(BASE_DIR, filename)

@app.route("/")
def home():
    return render_template("home.html")


@app.route("/recommendation")
def recommendation():
    return render_template("index.html")


@app.route("/inverse-soft-set")
def inverse_soft_set():
    return render_template("inverse_soft_set.html")


@app.route("/file/<path:filename>")
def file(filename):
    return send_from_directory("file", filename)


@app.route("/api/upload", methods=["POST"])
def upload():
    uploaded_file = request.files.get("file")

    if uploaded_file is None or uploaded_file.filename == "":
        return jsonify({
            "error": "Pilih file CSV atau Excel terlebih dahulu."
        }), 400

    extension = Path(uploaded_file.filename).suffix.lower()

    try:
        if extension == ".csv":
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin-1")

        elif extension in [".xlsx", ".xls"]:
            df = pd.read_excel(uploaded_file)

        else:
            return jsonify({
                "error": "Format file harus CSV, XLSX, atau XLS."
            }), 400

    except Exception as e:
        return jsonify({
            "error": f"File tidak dapat dibaca: {str(e)}"
        }), 400

    if df.empty:
        return jsonify({
            "error": "File tidak memiliki data."
        }), 400

    if len(df.columns) < 2:
        return jsonify({
            "error": (
                "Dataset harus memiliki minimal 2 kolom: "
                "kolom objek dan minimal 1 parameter."
            )
        }), 400

    if df.columns.duplicated().any():
        return jsonify({
            "error": "Nama kolom tidak boleh duplikat."
        }), 400

    if any(str(column).strip() == "" for column in df.columns):
        return jsonify({
            "error": "Semua kolom harus memiliki nama."
        }), 400

    df.columns = [str(column).strip() for column in df.columns]
    object_col = df.columns[0]
    parameters = []

    for column in df.columns[1:]:
        series = df[column]
        numeric_series = pd.to_numeric(series, errors="coerce")
        non_empty = series.notna() & series.astype(str).str.strip().ne("")

        is_numeric = (
            non_empty.sum() == 0
            or numeric_series[non_empty].notna().all()
        )

        parameter = {
            "name": column,
            "type": "numeric" if is_numeric else "categorical"
        }

        if is_numeric:
            df[column] = numeric_series
            valid_values = numeric_series.dropna()

            parameter["min"] = (
                float(valid_values.min())
                if not valid_values.empty
                else 0.0
            )

            parameter["max"] = (
                float(valid_values.max())
                if not valid_values.empty
                else 0.0
            )
        else:
            parameter["values"] = (
                series.dropna()
                .astype(str)
                .unique()
                .tolist()
            )

        parameters.append(parameter)

    objects = df[object_col].dropna().astype(str).tolist()

    return jsonify({
        "object_column": object_col,
        "objects": objects,
        "parameters": parameters,
        "rows": len(df),
        "columns": len(df.columns) - 1,
        "missing": int(df.isna().sum().sum()),
        "data": df.fillna("").astype(str).to_dict(orient="records")
    })


@app.route("/api/recommend", methods=["POST"])
def recommend():
    payload = request.get_json(silent=True) or {}
    data = payload.get("data", [])
    object_col = payload.get("object_column")
    rules = payload.get("rules", {})
    preferences = payload.get("preferences", [])

    if not data or not object_col:
        return jsonify({
            "error": "Dataset belum tersedia."
        }), 400

    if not preferences:
        return jsonify({
            "error": "Pilih minimal satu parameter preferensi."
        }), 400

    df = pd.DataFrame(data)

    for parameter, rule in rules.items():
        if parameter not in df.columns:
            return jsonify({
                "error": f"Parameter {parameter} tidak ditemukan."
            }), 400

        if rule["type"] == "numeric":
            df[parameter] = pd.to_numeric(
                df[parameter],
                errors="coerce"
            )

    soft_set = build_soft_set(df, rules, object_col)
    objects = df[object_col].dropna().astype(str).tolist()
    result = []

    for object_name in objects:
        matched = []

        for parameter in preferences:
            if object_name in soft_set.get(parameter, []):
                matched.append(parameter)

        score = len(matched) / len(preferences)

        result.append({
            "Nama": object_name,
            "Score": score,
            "Parameter Terpenuhi": matched
        })

    result.sort(key=lambda x: x["Score"], reverse=True)

    for index, item in enumerate(result, start=1):
        item["Peringkat"] = index

    max_score = max(
        (item["Score"] for item in result),
        default=0
    )

    best = [
        item for item in result
        if item["Score"] == max_score
    ]

    return jsonify({
        "soft_set": soft_set,
        "result": result,
        "best": best,
        "max_score": max_score
    })


if __name__ == "__main__":
    app.run(debug=True)