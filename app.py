from flask import Flask, request, jsonify, render_template, Response, send_file
from pymongo import MongoClient
from bson.objectid import ObjectId
from config import Config
import random
import json
import csv
import io
import re
from openpyxl import Workbook


app = Flask(__name__)
app.config.from_object(Config)


# ============================================================
# MONGODB CONNECTION
# ============================================================

MONGO_URI = app.config.get("MONGO_URI")
MONGO_DATABASE = app.config.get("MONGO_DATABASE")

if not MONGO_URI:
    print("WARNING: MONGO_URI is not configured.")

if not MONGO_DATABASE:
    print("WARNING: MONGO_DATABASE is not configured.")

client = None
db = None
dataset_collection = None

try:
    if MONGO_URI:
        client = MongoClient(
            MONGO_URI,
            serverSelectionTimeoutMS=5000
        )

        # Test connection
        client.admin.command("ping")

        db = client[MONGO_DATABASE]
        dataset_collection = db["datasets"]

        print("MongoDB connected successfully.")

except Exception as e:
    print("MongoDB connection error:", str(e))
    client = None
    db = None
    dataset_collection = None


# ============================================================
# HOME
# ============================================================

@app.route("/")
def home():
    return render_template("index.html")


# ============================================================
# GENERATE DATASET
# ============================================================

@app.route("/generate", methods=["POST"])
def generate_dataset():

    try:

        # ----------------------------------------------------
        # CHECK REQUEST
        # ----------------------------------------------------

        data = request.get_json(silent=True)

        if not data:
            return jsonify({
                "success": False,
                "error": "Invalid JSON request."
            }), 400


        # ----------------------------------------------------
        # GET PROMPT
        # ----------------------------------------------------

        prompt = str(
            data.get("prompt", "")
        ).strip()

        if not prompt:
            return jsonify({
                "success": False,
                "error": "Prompt is required."
            }), 400


        prompt_lower = prompt.lower()


        # ----------------------------------------------------
        # NUMBER OF RECORDS
        # ----------------------------------------------------

        number_of_records = 5

        patterns = [
            r"\b(\d+)\s+(?:records?|rows?|entries?)\b",

            r"\b(?:generate|create|make|give|show)\s+(\d+)\b",

            r"\b(\d+)\s+(?:students?|employees?|customers?|products?|users?)\b"
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                prompt_lower
            )

            if match:
                number_of_records = int(
                    match.group(1)
                )
                break


        if number_of_records < 1:
            number_of_records = 1


        if number_of_records > 1000:
            number_of_records = 1000


        # ----------------------------------------------------
        # DATASET TYPE
        # ----------------------------------------------------

        if "student" in prompt_lower:

            dataset_type = "student"

        elif "employee" in prompt_lower:

            dataset_type = "employee"

        elif "product" in prompt_lower:

            dataset_type = "product"

        elif "customer" in prompt_lower:

            dataset_type = "customer"

        else:

            dataset_type = "general"


        # ----------------------------------------------------
        # COLUMN ALIASES
        # ----------------------------------------------------

        column_aliases = {

            "name": [
                "name",
                "full name"
            ],

            "age": [
                "age"
            ],

            "city": [
                "city",
                "location"
            ],

            "email": [
                "email",
                "email address"
            ],

            "phone": [
                "phone",
                "phone number",
                "mobile"
            ],

            "gender": [
                "gender"
            ],

            "occupation": [
                "occupation",
                "job"
            ],

            "department": [
                "department",
                "dept"
            ],

            "salary": [
                "salary"
            ],

            "income": [
                "income"
            ],

            "experience": [
                "experience",
                "years of experience"
            ],

            "branch": [
                "branch",
                "stream"
            ],

            "cgpa": [
                "cgpa",
                "gpa"
            ],

            "grade": [
                "grade"
            ],

            "product": [
                "product",
                "product name",
                "item"
            ],

            "price": [
                "price",
                "cost"
            ],

            "quantity": [
                "quantity",
                "qty"
            ],

            "company": [
                "company",
                "organization",
                "organisation"
            ]
        }


        # ----------------------------------------------------
        # SELECT COLUMNS FROM PROMPT
        # ----------------------------------------------------

        selected_columns = []

        for column, aliases in column_aliases.items():

            for alias in aliases:

                if alias in prompt_lower:

                    selected_columns.append(
                        column
                    )

                    break


        selected_columns = list(
            dict.fromkeys(
                selected_columns
            )
        )


        # ----------------------------------------------------
        # DEFAULT COLUMNS
        # ----------------------------------------------------

        if not selected_columns:

            if dataset_type == "student":

                selected_columns = [
                    "name",
                    "age",
                    "branch",
                    "cgpa"
                ]

            elif dataset_type == "employee":

                selected_columns = [
                    "name",
                    "age",
                    "department",
                    "salary"
                ]

            elif dataset_type == "product":

                selected_columns = [
                    "product",
                    "price",
                    "quantity"
                ]

            elif dataset_type == "customer":

                selected_columns = [
                    "name",
                    "age",
                    "city",
                    "occupation",
                    "income"
                ]

            else:

                selected_columns = [
                    "name",
                    "age",
                    "city",
                    "occupation",
                    "income"
                ]


        # ----------------------------------------------------
        # DATA VALUES
        # ----------------------------------------------------

        names = [
            "Aarav",
            "Rahul",
            "Arjun",
            "Zoya",
            "Aisha",
            "Kabir",
            "Sara",
            "Aditya",
            "Riya",
            "Imran",
            "Priya",
            "Omar",
            "Ananya",
            "Vikram"
        ]


        cities = [
            "Delhi",
            "Mumbai",
            "Hyderabad",
            "Bangalore",
            "Kashmir",
            "Chennai",
            "Pune",
            "Jaipur",
            "Kolkata",
            "Ahmedabad"
        ]


        occupations = [
            "Engineer",
            "Teacher",
            "Designer",
            "Manager",
            "Developer",
            "Analyst",
            "Doctor",
            "Accountant"
        ]


        departments = [
            "IT",
            "HR",
            "Finance",
            "Sales",
            "Marketing",
            "Operations",
            "Engineering"
        ]


        branches = [
            "CSE",
            "ECE",
            "EEE",
            "ME",
            "Civil",
            "IT",
            "AI",
            "DS"
        ]


        products = [
            "Laptop",
            "Smartphone",
            "Headphones",
            "Keyboard",
            "Mouse",
            "Monitor",
            "Tablet",
            "Smart Watch"
        ]


        companies = [
            "Google",
            "Microsoft",
            "Amazon",
            "Infosys",
            "TCS",
            "Accenture",
            "IBM",
            "Wipro"
        ]


        # ----------------------------------------------------
        # GENERATE DATASET
        # ----------------------------------------------------

        dataset = []


        for i in range(
            number_of_records
        ):

            record = {}


            for column in selected_columns:

                if column == "name":

                    value = random.choice(
                        names
                    )


                elif column == "age":

                    value = random.randint(
                        18,
                        60
                    )


                elif column == "city":

                    value = random.choice(
                        cities
                    )


                elif column == "email":

                    name = random.choice(
                        names
                    ).lower()

                    number = random.randint(
                        10,
                        9999
                    )

                    value = (
                        f"{name}"
                        f"{number}"
                        "@example.com"
                    )


                elif column == "phone":

                    value = (
                        "9"
                        +
                        "".join(
                            str(
                                random.randint(
                                    0,
                                    9
                                )
                            )
                            for _ in range(9)
                        )
                    )


                elif column == "gender":

                    value = random.choice([
                        "Male",
                        "Female"
                    ])


                elif column == "occupation":

                    value = random.choice(
                        occupations
                    )


                elif column == "department":

                    value = random.choice(
                        departments
                    )


                elif column == "salary":

                    value = random.randint(
                        30000,
                        150000
                    )


                elif column == "income":

                    value = random.randint(
                        25000,
                        150000
                    )


                elif column == "experience":

                    value = random.randint(
                        0,
                        20
                    )


                elif column == "branch":

                    value = random.choice(
                        branches
                    )


                elif column == "cgpa":

                    value = round(
                        random.uniform(
                            6.0,
                            10.0
                        ),
                        2
                    )


                elif column == "grade":

                    value = random.choice([
                        "A",
                        "B",
                        "C",
                        "D"
                    ])


                elif column == "product":

                    value = random.choice(
                        products
                    )


                elif column == "price":

                    value = random.randint(
                        500,
                        100000
                    )


                elif column == "quantity":

                    value = random.randint(
                        1,
                        20
                    )


                elif column == "company":

                    value = random.choice(
                        companies
                    )


                else:

                    value = "N/A"


                record[column] = value


            dataset.append(record)


        # ----------------------------------------------------
        # SAVE TO MONGODB
        # ----------------------------------------------------

        document = {
            "prompt": prompt,
            "dataset": dataset
        }


        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": (
                    "MongoDB is not connected. "
                    "Please check MONGO_URI and "
                    "MONGO_DATABASE environment variables."
                )
            }), 500


        result = dataset_collection.insert_one(
            document
        )


        # ----------------------------------------------------
        # SUCCESS RESPONSE
        # ----------------------------------------------------

        return jsonify({

            "success": True,

            "message":
                "Dataset generated and saved successfully",

            "dataset_id":
                str(result.inserted_id),

            "rows":
                number_of_records,

            "columns":
                selected_columns,

            "dataset":
                dataset

        })


    # ========================================================
    # ERROR HANDLING
    # ========================================================

    except Exception as e:

        print(
            "ERROR IN /generate:",
            str(e)
        )

        return jsonify({

            "success": False,

            "error":
                f"Dataset generation failed: {str(e)}"

        }), 500


# ============================================================
# GET ALL DATASETS
# ============================================================

@app.route("/datasets", methods=["GET"])
def get_datasets():

    try:

        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": "MongoDB is not connected."
            }), 500


        datasets = list(
            dataset_collection.find()
        )


        for dataset in datasets:

            dataset["_id"] = str(
                dataset["_id"]
            )


        return jsonify({

            "success": True,

            "datasets":
                datasets

        })


    except Exception as e:

        print(
            "ERROR IN /datasets:",
            str(e)
        )

        return jsonify({

            "success": False,

            "error":
                f"Failed to load datasets: {str(e)}"

        }), 500


# ============================================================
# UPDATE DATASET
# ============================================================

@app.route(
    "/dataset/<dataset_id>",
    methods=["PUT"]
)
def update_dataset(dataset_id):

    try:

        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": "MongoDB is not connected."
            }), 500


        data = request.get_json(
            silent=True
        )


        if not data:

            return jsonify({
                "success": False,
                "error": "Invalid JSON request."
            }), 400


        updated_dataset = data.get(
            "dataset"
        )


        if updated_dataset is None:

            return jsonify({
                "success": False,
                "error": "Dataset is required"
            }), 400


        try:

            object_id = ObjectId(
                dataset_id
            )

        except Exception:

            return jsonify({
                "success": False,
                "error": "Invalid dataset ID"
            }), 400


        result = dataset_collection.update_one(

            {
                "_id": object_id
            },

            {
                "$set": {
                    "dataset":
                        updated_dataset
                }
            }

        )


        if result.matched_count == 0:

            return jsonify({
                "success": False,
                "error": "Dataset not found"
            }), 404


        return jsonify({

            "success": True,

            "message":
                "Dataset updated successfully"

        })


    except Exception as e:

        print(
            "ERROR IN UPDATE:",
            str(e)
        )

        return jsonify({

            "success": False,

            "error":
                f"Failed to update dataset: {str(e)}"

        }), 500


# ============================================================
# DELETE DATASET
# ============================================================

@app.route(
    "/dataset/<dataset_id>",
    methods=["DELETE"]
)
def delete_dataset(dataset_id):

    try:

        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": "MongoDB is not connected."
            }), 500


        try:

            object_id = ObjectId(
                dataset_id
            )

        except Exception:

            return jsonify({
                "success": False,
                "error": "Invalid dataset ID"
            }), 400


        result = dataset_collection.delete_one({

            "_id": object_id

        })


        if result.deleted_count == 0:

            return jsonify({
                "success": False,
                "error": "Dataset not found"
            }), 404


        return jsonify({

            "success": True,

            "message":
                "Dataset deleted successfully"

        })


    except Exception as e:

        print(
            "ERROR IN DELETE:",
            str(e)
        )

        return jsonify({

            "success": False,

            "error":
                f"Failed to delete dataset: {str(e)}"

        }), 500


# ============================================================
# DOWNLOAD JSON
# ============================================================

@app.route(
    "/dataset/<dataset_id>/download/json",
    methods=["GET"]
)
def download_json(dataset_id):

    try:

        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": "MongoDB is not connected."
            }), 500


        try:

            object_id = ObjectId(
                dataset_id
            )

        except Exception:

            return jsonify({
                "success": False,
                "error": "Invalid dataset ID"
            }), 400


        dataset_document = (
            dataset_collection.find_one({
                "_id": object_id
            })
        )


        if not dataset_document:

            return jsonify({
                "success": False,
                "error": "Dataset not found"
            }), 404


        dataset = dataset_document.get(
            "dataset",
            []
        )


        json_data = json.dumps(
            dataset,
            indent=4
        )


        response = Response(
            json_data,
            mimetype="application/json"
        )


        response.headers[
            "Content-Disposition"
        ] = (
            "attachment; "
            f"filename=dataset_{dataset_id}.json"
        )


        return response


    except Exception as e:

        return jsonify({
            "success": False,
            "error":
                f"Failed to download JSON: {str(e)}"
        }), 500


# ============================================================
# DOWNLOAD CSV
# ============================================================

@app.route(
    "/dataset/<dataset_id>/download/csv",
    methods=["GET"]
)
def download_csv(dataset_id):

    try:

        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": "MongoDB is not connected."
            }), 500


        try:

            object_id = ObjectId(
                dataset_id
            )

        except Exception:

            return jsonify({
                "success": False,
                "error": "Invalid dataset ID"
            }), 400


        dataset_document = (
            dataset_collection.find_one({
                "_id": object_id
            })
        )


        if not dataset_document:

            return jsonify({
                "success": False,
                "error": "Dataset not found"
            }), 404


        dataset = dataset_document.get(
            "dataset",
            []
        )


        if not dataset:

            return jsonify({
                "success": False,
                "error": "Dataset is empty"
            }), 400


        output = io.StringIO()


        fieldnames = list(
            dataset[0].keys()
        )


        writer = csv.DictWriter(
            output,
            fieldnames=fieldnames
        )


        writer.writeheader()

        writer.writerows(
            dataset
        )


        response = Response(
            output.getvalue(),
            mimetype="text/csv"
        )


        response.headers[
            "Content-Disposition"
        ] = (
            "attachment; "
            f"filename=dataset_{dataset_id}.csv"
        )


        return response


    except Exception as e:

        return jsonify({
            "success": False,
            "error":
                f"Failed to download CSV: {str(e)}"
        }), 500


# ============================================================
# DOWNLOAD EXCEL
# ============================================================

@app.route(
    "/dataset/<dataset_id>/download/excel",
    methods=["GET"]
)
def download_excel(dataset_id):

    try:

        if dataset_collection is None:

            return jsonify({
                "success": False,
                "error": "MongoDB is not connected."
            }), 500


        try:

            object_id = ObjectId(
                dataset_id
            )

        except Exception:

            return jsonify({
                "success": False,
                "error": "Invalid dataset ID"
            }), 400


        dataset_document = (
            dataset_collection.find_one({
                "_id": object_id
            })
        )


        if not dataset_document:

            return jsonify({
                "success": False,
                "error": "Dataset not found"
            }), 404


        dataset = dataset_document.get(
            "dataset",
            []
        )


        if not dataset:

            return jsonify({
                "success": False,
                "error": "Dataset is empty"
            }), 400


        workbook = Workbook()


        worksheet = workbook.active


        worksheet.title = "Dataset"


        fieldnames = list(
            dataset[0].keys()
        )


        worksheet.append(
            fieldnames
        )


        for row in dataset:

            worksheet.append([

                row.get(field)

                for field in fieldnames

            ])


        output = io.BytesIO()


        workbook.save(
            output
        )


        output.seek(0)


        return send_file(

            output,

            as_attachment=True,

            download_name=
                f"dataset_{dataset_id}.xlsx",

            mimetype=
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )


    except Exception as e:

        return jsonify({
            "success": False,
            "error":
                f"Failed to download Excel: {str(e)}"
        }), 500


# ============================================================
# GLOBAL JSON ERROR HANDLERS
# ============================================================

@app.errorhandler(404)
def handle_404(error):

    return jsonify({

        "success": False,

        "error":
            "API endpoint not found."

    }), 404


@app.errorhandler(405)
def handle_405(error):

    return jsonify({

        "success": False,

        "error":
            "HTTP method not allowed."

    }), 405


@app.errorhandler(500)
def handle_500(error):

    print(
        "GLOBAL SERVER ERROR:",
        str(error)
    )

    return jsonify({

        "success": False,

        "error":
            "Internal server error."

    }), 500


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )