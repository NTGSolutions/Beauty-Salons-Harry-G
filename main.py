from fastapi import FastAPI, Request
from arcgis.gis import GIS
from arcgis.features import FeatureLayer
import os
import qrcode
from datetime import datetime

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# =========================
# APP INIT
# =========================
app = FastAPI()

# =========================
# AGOL AUTH
# =========================
AGOL_USERNAME = os.getenv("AGOL_USERNAME")
AGOL_PASSWORD = os.getenv("AGOL_PASSWORD")

if not AGOL_USERNAME or not AGOL_PASSWORD:
    raise Exception("AGOL credentials not set in environment variables")

gis = GIS("https://www.arcgis.com", AGOL_USERNAME, AGOL_PASSWORD)

# =========================
# FEATURE LAYER
# =========================
SURVEY_LAYER_URL = "https://services6.arcgis.com/345WScIubRHps95b/arcgis/rest/services/service_0f5bfa7121a34359a1b4a1402559cb1f/FeatureServer/0"
layer = FeatureLayer(SURVEY_LAYER_URL, gis=gis)

# =========================
# TEMPLATE
# =========================
TEMPLATE_PATH = "Salons_inspection_form_report.docx"

if not os.path.exists(TEMPLATE_PATH):
    raise Exception(f"{TEMPLATE_PATH} not found in project root")

# =========================
# TEMP PAYLOAD STORAGE
# =========================
LAST_PAYLOAD = {}
LAST_ERROR = None

# =========================
# HEALTH CHECK
# =========================
@app.get("/")
def home():
    return {"status": "running"}

# =========================
# DEBUG ENDPOINT
# =========================
@app.get("/debug")
def debug():
    return {
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "username_set": bool(AGOL_USERNAME),
        "password_set": bool(AGOL_PASSWORD),
        "layer_url": SURVEY_LAYER_URL
    }

# =========================
# LAST PAYLOAD
# =========================
@app.get("/last-payload")
def last_payload():
    return {
        "last_error": LAST_ERROR,
        "payload": LAST_PAYLOAD
    }

# =========================
# TEST QUERY
# =========================
@app.get("/test-query/{objectid}")
def test_query(objectid: int):
    result = layer.query(where=f"OBJECTID={objectid}", out_fields="*")
    return {
        "found": len(result.features),
        "attributes": result.features[0].attributes if result.features else None
    }

# =========================
# TEST UPDATE
# =========================
@app.get("/test-update/{objectid}")
def test_update(objectid: int):
    result = layer.edit_features(updates=[{
        "attributes": {
            "OBJECTID": objectid,
            "report_status": "test_ok",
            "report_url": "https://example.com/test.docx"
        }
    }])
    return {"edit_result": result}

# =========================
# HELPER: EXTRACT OBJECTID
# =========================
def extract_objectid(payload):
    if "submittedRecord" in payload:
        attrs = payload["submittedRecord"].get("attributes", {})
        if "OBJECTID" in attrs:
            return attrs["OBJECTID"]

    if "serverResponse" in payload:
        sr = payload["serverResponse"]
        if isinstance(sr, dict):
            if "objectId" in sr:
                return sr["objectId"]
            if "editResults" in sr and sr["editResults"]:
                first = sr["editResults"][0]
                if "objectId" in first:
                    return first["objectId"]

    if "feature" in payload:
        feature = payload["feature"]
        if isinstance(feature, dict):
            attrs = feature.get("attributes", {})
            if "OBJECTID" in attrs:
                return attrs["OBJECTID"]
            result = feature.get("result", {})
            if "objectId" in result:
                return result["objectId"]

    if "features" in payload and payload["features"]:
        first = payload["features"][0]
        attrs = first.get("attributes", {})
        if "OBJECTID" in attrs:
            return attrs["OBJECTID"]

    if isinstance(payload, dict):
        for key, value in payload.items():
            if key in ("OBJECTID", "objectId"):
                return value
            found = extract_objectid(value)
            if found is not None:
                return found

    if isinstance(payload, list):
        for item in payload:
            found = extract_objectid(item)
            if found is not None:
                return found

    return None

# =========================
# QR GENERATOR
# =========================
def generate_qr(url, path):
    img = qrcode.make(url)
    img.save(path)

# =========================
# UPLOAD REPORT TO AGOL
# =========================
def upload_report_to_agol(file_path, objectid):
    root_folder = gis.content.folders.get()

    item_properties = {
        "title": f"Report_{objectid}",
        "type": "Microsoft Word",
        "tags": ["survey123", "report", "automation"],
        "snippet": f"Automatically generated report for Survey123 submission {objectid}"
    }

    report_item = root_folder.add(
        item_properties=item_properties,
        file=file_path
    ).result()

    report_item.sharing.sharing_level = "EVERYONE"

    return f"https://www.arcgis.com/home/item.html?id={report_item.itemid}"

# =========================
# REPORT GENERATION
# =========================
def generate_report(attributes, objectid):
    os.makedirs("output", exist_ok=True)

    docx_file = os.path.join("output", f"report_{objectid}.docx")
    qr_file = os.path.join("output", f"qr_{objectid}.png")

    # Temporary QR target for first render
    temp_url = f"https://www.arcgis.com/home/item.html?id=temp-{objectid}"
    generate_qr(temp_url, qr_file)

    edit_date = attributes.get("EditDate")
    if edit_date:
        edit_date = datetime.fromtimestamp(edit_date / 1000).strftime("%Y-%m-%d %H:%M:%S")
    else:
        edit_date = "N/A"

    doc = DocxTemplate(TEMPLATE_PATH)
    qr_image = InlineImage(doc, qr_file, width=Mm(25))

    context = {
        "name_premise": attributes.get("name_premise", "N/A"),
        "owner_name": attributes.get("owner_name", "N/A"),
        "address__": attributes.get("address__", "N/A"),
        "Surname": attributes.get("Surname", "N/A"),
        "ID_Num": attributes.get("ID_Num", "N/A"),
        "contact_person": attributes.get("contact_person", "N/A"),
        "inspection_date": edit_date,
        "telephone_no": attributes.get("telephone_no", "N/A"),
         "Description": attributes.get("Description", "N/A"),
         "premise_type": attributes.get("premise_type", "N/A"),

        "pes__": attributes.get("pes__", "N/A"),
        "pest_control": attributes.get("pest_control", "N/A"),
        "waste_disposal": attributes.get("waste_disposal", "N/A"),
        "waste_disposal_": attributes.get("waste_disposal_", "N/A"),
        "female": attributes.get("female", "N/A"),
        "male": attributes.get("male", "N/A"),
        "health_certificate": attributes.get("health_certificate", "N/A"),
        "expiry_date": attributes.get("expiry_date", "N/A"),
        "comment_coa": attributes.get("comment_coa", "N/A"),
        "relevant_permit": attributes.get("relevant_permit", "N/A"),

        "comment1": attributes.get("comment1", "N/A"),
        "collection_healthcare": attributes.get("collection_healthcare", "N/A"),
        "comm1": attributes.get("comm1", "N/A"),
        "training": attributes.get("training", "N/A"),
        "municname": attributes.get("municname", "N/A"),
        "comm2": attributes.get("comm2", "N/A"),
        "comme3": attributes.get("comme3", "N/A"),
        "incident_reg": attributes.get("incident_reg", "N/A"),
        "training_conducted": attributes.get("training_conducted", "N/A"),
        "comme4": attributes.get("comme4", "N/A"),

        "internal_walls": attributes.get("internal_walls", "N/A"),
        "cleanable_material": attributes.get("cleanable_material", "N/A"),
        "comment2": attributes.get("comment2", "N/A"),
        "coloured_paint": attributes.get("coloured_paint", "N/A"),
        "comment3": attributes.get("comment3", "N/A"),
        "floors_": attributes.get("floors_", "N/A"),
        "smooth": attributes.get("smooth", "N/A"),
        "comment4": attributes.get("comment4", "N/A"),
        "cleanable_material_1": attributes.get("cleanable_material_1", "N/A"),

        "comment5": attributes.get("comment5", "N/A"),
        "ceiling0": attributes.get("ceiling0", "N/A"),
        "dustproof": attributes.get("dustproof", "N/A"),
        "comment6": attributes.get("comment6", "N/A"),
        "employees0": attributes.get("employees0", "N/A"),
        "comment7": attributes.get("comment7", "N/A"),
        "employees1": attributes.get("employees1", "N/A"),
        "comment8": attributes.get("comment8", "N/A"),

        "different_sexes": attributes.get("different_sexes", "N/A"),
        "comment9": attributes.get("comment9", "N/A"),
        "hot_and_cold": attributes.get("hot_and_cold", "N/A"),
        "comment10": attributes.get("comment10", "N/A"),
        "working_surface": attributes.get("working_surface", "N/A"),
        "comment11": attributes.get("comment11", "N/A"),
        "durable": attributes.get("durable", "N/A"),
        "comment12": attributes.get("comment12", "N/A"),

        "cleanable_mat1": attributes.get("cleanable_mat1", "N/A"),
        "comment13": attributes.get("comment13", "N/A"),
        "adequate_cupboards": attributes.get("adequate_cupboards", "N/A"),
        "comment14": attributes.get("comment14", "N/A"),
        "personal_belongings": attributes.get("personal_belongings", "N/A"),
        "comment15": attributes.get("comment15", "N/A"),
        "hand_basin": attributes.get("hand_basin", "N/A"),
        "comment16": attributes.get("comment16", "N/A"),

        "hot_cold": attributes.get("hot_cold", "N/A"),
        "comment17": attributes.get("comment17", "N/A"),
        "adequate_refuse": attributes.get("adequate_refuse", "N/A"),
        "comment18": attributes.get("comment18", "N/A"),
        "clean_sanitary": attributes.get("clean_sanitary", "N/A"),
        "comment19": attributes.get("comment19", "N/A"),
        "instruments_disinfected": attributes.get("instruments_disinfected", "N/A"),
        "comment20": attributes.get("comment20", "N/A"),

        "type_disinfectant": attributes.get("type_disinfectant", "N/A"),
        "comment21": attributes.get("comment21", "N/A"),
        "blood_sterilizer": attributes.get("blood_sterilizer", "N/A"),
        "comment22": attributes.get("comment22", "N/A"),
        "type_sterilizer": attributes.get("type_sterilizer", "N/A"),
        "comment23": attributes.get("comment23", "N/A"),
        "waterproof_apron": attributes.get("waterproof_apron", "N/A"),
        "comment24": attributes.get("comment24", "N/A"),

        "gloves": attributes.get("gloves", "N/A"),
        "comment25": attributes.get("comment25", "N/A"),
        "dust_mask": attributes.get("dust_mask", "N/A"),
        "comment26": attributes.get("comment26", "N/A"),
        "plastic_": attributes.get("plastic_", "N/A"),
        "comment27": attributes.get("comment27", "N/A"),
        "disposable": attributes.get("disposable", "N/A"),
        "comment28": attributes.get("comment28", "N/A"),

        "adequate_towels": attributes.get("adequate_towels", "N/A"),
        "comment29": attributes.get("comment29", "N/A"),
        "laundry": attributes.get("laundry", "N/A"),
        "comment30": attributes.get("comment30", "N/A"),
        "no_animals": attributes.get("no_animals", "N/A"),
        "comment31": attributes.get("comment31", "N/A"),
        "beverages": attributes.get("beverages", "N/A"),
        "comment32": attributes.get("comment32", "N/A"),
        "single_bowl": attributes.get("single_bowl", "N/A"),
        "comment33": attributes.get("comment33", "N/A"),
        "no_food": attributes.get("no_food", "N/A"),
        "comment34": attributes.get("comment34", "N/A"),
        "no_sleeping": attributes.get("no_sleeping", "N/A"),
        "comment35": attributes.get("comment35", "N/A"),

        "wash_management": attributes.get("wash_management", "N/A"),
        "comment36": attributes.get("comment36", "N/A"),
        "labeleld_health": attributes.get("labeleld_health", "N/A"),
        "comment37": attributes.get("comment37", "N/A"),
        "health_care01": attributes.get("health_care01", "N/A"),
        "comment38": attributes.get("comment38", "N/A"),
        "disposed": attributes.get("disposed", "N/A"),
        "comment39": attributes.get("comment39", "N/A"),

        "must_contain": attributes.get("must_contain", "N/A"),
        "comment40": attributes.get("comment40", "N/A"),
        "clearly_labelled": attributes.get("clearly_labelled", "N/A"),
        "comment41": attributes.get("comment41", "N/A"),
        "health_care_risk": attributes.get("health_care_risk", "N/A"),
        "comment42": attributes.get("comment42", "N/A"),

        "isnpect_records": attributes.get("isnpect_records", "N/A"),
        "comment43": attributes.get("comment43", "N/A"),
        "health_info": attributes.get("health_info", "N/A"),
        "comment44": attributes.get("comment44", "N/A"),
        "easily_accessible": attributes.get("easily_accessible", "N/A"),
        "comment45": attributes.get("comment45", "N/A"),

        "protective_eye_wear": attributes.get("protective_eye_wear", "N/A"),
        "comment46": attributes.get("comment46", "N/A"),
        "disinfectant_used": attributes.get("disinfectant_used", "N/A"),
        "comment47": attributes.get("comment47", "N/A"),
        "disinfectant1": attributes.get("disinfectant1", "N/A"),
        "comment48": attributes.get("comment48", "N/A"),

        "records_kept": attributes.get("records_kept", "N/A"),
        "comment49": attributes.get("comment49", "N/A"),
        "non_toxic": attributes.get("non_toxic", "N/A"),
        "comment50": attributes.get("comment50", "N/A"),
        "sterile_containers": attributes.get("sterile_containers", "N/A"),
        "comment51": attributes.get("comment51", "N/A"),

        "stencil_disinfected": attributes.get("stencil_disinfected", "N/A"),
        "comment52": attributes.get("comment52", "N/A"),
        "stencils_discarded": attributes.get("stencils_discarded", "N/A"),
        "comment53": attributes.get("comment53", "N/A"),
        "customers_body": attributes.get("customers_body", "N/A"),
        "comment54": attributes.get("comment54", "N/A"),

        "waterproof_aprons": attributes.get("waterproof_aprons", "N/A"),
        "comment55": attributes.get("comment55", "N/A"),
        "soap_availability": attributes.get("soap_availability", "N/A"),
        "comment56": attributes.get("comment56", "N/A"),
        "storage_facility": attributes.get("storage_facility", "N/A"),
        "comment57": attributes.get("comment57", "N/A"),
        "sterilizer_used": attributes.get("sterilizer_used", "N/A"),
        "comment58": attributes.get("comment58", "N/A"),
        "good_sanitary": attributes.get("good_sanitary", "N/A"),
        "comment59": attributes.get("comment59", "N/A"),
        "recommendations2": attributes.get("recommendations2", "N/A"),
        "recomm": attributes.get("recomm", "N/A"),
        "recommendations003": attributes.get("recommendations003", "N/A"),
        "other_action_taken": attributes.get("other_action_taken", "N/A"),
        "risk_rating": attributes.get("risk_rating", "N/A"),

        "recommedations_": attributes.get("recommedations_", "N/A"),
        "compliance": attributes.get("compliance", "N/A"),
        "additional_pictures": attributes.get("additional_pictures", "N/A"),
        "EHP": attributes.get("EHP", "N/A"),
        "email_address": attributes.get("email_address", "N/A"),
        "contacts": attributes.get("contacts", "N/A"),
        "signature": attributes.get("signature", "N/A"),
        "pic_": attributes.get("pic_", "N/A"),
        "manager_signature": attributes.get("manager_signature", "N/A"),

        "qr_code": qr_image
    }

    doc.render(context)
    doc.save(docx_file)

    real_url = upload_report_to_agol(docx_file, objectid)
    return real_url

# =========================
# UPDATE FEATURE
# =========================
def update_feature(objectid, url, status):
    result = layer.edit_features(updates=[{
        "attributes": {
            "OBJECTID": objectid,
            "report_url": url,
            "report_status": status
        }
    }])
    return result

# =========================
# WEBHOOK ENDPOINT
# =========================
@app.post("/webhook/survey123")
async def survey_webhook(request: Request):
    global LAST_PAYLOAD, LAST_ERROR

    payload = await request.json()
    LAST_PAYLOAD = payload
    LAST_ERROR = None
    objectid = None

    try:
        objectid = extract_objectid(payload)

        if objectid is None:
            LAST_ERROR = f"OBJECTID not found. Payload keys: {list(payload.keys()) if isinstance(payload, dict) else 'not a dict'}"
            return {
                "status": "failed",
                "error": LAST_ERROR
            }

        update_feature(objectid, "webhook_received", "received")

        result = layer.query(where=f"OBJECTID={objectid}", out_fields="*")

        if not result.features:
            update_feature(objectid, "query_failed", "failed")
            LAST_ERROR = f"No feature found for OBJECTID {objectid}"
            return {
                "status": "failed",
                "error": LAST_ERROR
            }

        attributes = result.features[0].attributes

        update_feature(objectid, "query_ok", "queried")

        report_url = generate_report(attributes, objectid)

        edit_result = update_feature(objectid, report_url, "completed")

        return {
            "status": "success",
            "objectid": objectid,
            "report_url": report_url,
            "edit_result": str(edit_result)
        }

    except Exception as e:
        LAST_ERROR = str(e)
        if objectid is not None:
            try:
                update_feature(objectid, f"ERROR: {str(e)}", "failed")
            except Exception:
                pass

        return {
            "status": "failed",
            "objectid": objectid,
            "error": str(e)
        }
