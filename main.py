import os
import re
import traceback
import datetime
from typing import Dict, Any, Optional, List, Tuple

from arcgis.gis import GIS
from arcgis.features import FeatureLayer
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from jinja2 import Environment

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse


# =========================================================
# FASTAPI APPLICATION
# =========================================================

app = FastAPI()


# =========================================================
# CONFIGURATION
# =========================================================

# ---------------------------------------------------------
# ArcGIS Online
# ---------------------------------------------------------

AGOL_URL = os.getenv(
    "AGOL_URL",
    "https://www.arcgis.com"
)

AGOL_USERNAME = os.getenv(
    "AGOL_USERNAME"
)

AGOL_PASSWORD = os.getenv(
    "AGOL_PASSWORD"
)


# ---------------------------------------------------------
# Feature Layer
# ---------------------------------------------------------
#
# IMPORTANT:
# This MUST point to the actual layer.
#
# Correct:
# .../FeatureServer/0
#
# Incorrect:
# .../FeatureServer
# ---------------------------------------------------------

FEATURE_LAYER_URL = os.getenv(
    "FEATURE_LAYER_URL"
)


# =========================================================
# FIELD NAMES
# =========================================================

INSPECTION_STATUS_FIELD = os.getenv(
    "INSPECTION_STATUS_FIELD",
    "inspection_status"
)


COMPLIANCE_FIELD = os.getenv(
    "COMPLIANCE_FIELD",
    "compliance_status"
)


ACTION_TAKEN_FIELD = os.getenv(
    "ACTION_TAKEN_FIELD",
    "action_taken"
)


# ---------------------------------------------------------
# Processing status field
# ---------------------------------------------------------
#
# This field is used to determine whether the report
# generation process has completed successfully.
# ---------------------------------------------------------

REPORT_STATUS_FIELD = os.getenv(
    "REPORT_STATUS_FIELD",
)


# =========================================================
# DOCUMENT TEMPLATES
# =========================================================

FINE_TEMPLATE_PATH = os.getenv(
    "FINE_TEMPLATE_PATH",
    "./templates/FINE_TEMPLATE_PATH.docx"
)


COMPLIANCE_NOTICE_TEMPLATE_PATH = os.getenv(
    "COMPLIANCE_NOTICE_TEMPLATE_PATH",
    "./templates/COMPLIANCE_NOTICE_TEMPLATE_PATH.docx"
)


PROHIBITION_NOTICE_TEMPLATE_PATH = os.getenv(
    "PROHIBITION_NOTICE_TEMPLATE_PATH",
    "./templates/PROHIBITION_NOTICE_TEMPLATE_PATH.docx"
)

REPORT_TEMPLATE_PATH = os.getenv(
    "REPORT_TEMPLATE_PATH",
    "./templates/REPORT_TEMPLATE_PATH.docx"
)


# =========================================================
# TEMPORARY DIRECTORY
# =========================================================

TEMP_DIR = os.getenv(
    "TEMP_DIR",
    "/tmp/reports"
)

os.makedirs(
    TEMP_DIR,
    exist_ok=True
)


# =========================================================
# PROCESSING OPTIONS
# =========================================================

SKIP_IF_ATTACHMENT_ALREADY_EXISTS = (
    os.getenv(
        "SKIP_IF_ATTACHMENT_ALREADY_EXISTS",
        "false"
    ).lower() == "true"
)


# =========================================================
# IMAGE / SIGNATURE CONFIGURATION
# =========================================================

IMAGE_FIELD_NAMES = [
    "Ehp_signature",
    "manager_signature",
    "SIGNATURE_OF_PERSON_IN_CHARGE_",
]

SIGNATURE_IMAGE_WIDTH_MM = float(
    os.getenv("SIGNATURE_IMAGE_WIDTH_MM", "35")
)

REQUIRE_SIGNATURE_IMAGES = (
    os.getenv("REQUIRE_SIGNATURE_IMAGES", "false").lower() == "true"
)


# =========================================================
# BASIC HELPERS
# =========================================================

def now_stamp() -> str:
    """
    Creates a timestamp for generated filenames.
    """

    return datetime.datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )


def out_docx_path(
    object_id: Any,
    label: str = "report"
) -> str:
    """
    Creates the temporary path for a generated DOCX.
    """

    safe_oid = (
        str(object_id)
        .replace("/", "_")
        .replace("\\", "_")
    )

    return os.path.join(
        TEMP_DIR,
        f"{label}_{safe_oid}_{now_stamp()}.docx"
    )


def require_env(name: str) -> str:
    """
    Ensures that a required environment variable exists.
    """

    value = os.getenv(name)

    if not value:
        raise RuntimeError(
            f"Missing required environment variable: {name}"
        )

    return value


# =========================================================
# TEMPLATE VALIDATION
# =========================================================

def validate_paths() -> None:
    """
    Confirms that all Word templates exist.
    """

    templates = {
        "FINE": FINE_TEMPLATE_PATH,
        "COMPLIANCE NOTICE": COMPLIANCE_NOTICE_TEMPLATE_PATH,
        "PROHIBITION NOTICE": PROHIBITION_NOTICE_TEMPLATE_PATH,
        "REPORT TEMPLATE": REPORT_TEMPLATE_PATH,
    }

    missing = []

    for label, path in templates.items():

        if not os.path.exists(path):

            missing.append(
                f"{label} template not found: {path}"
            )

    if missing:

        raise FileNotFoundError(
            " | ".join(missing)
        )

    print(
        "All document templates were found."
    )


# =========================================================
# NORMALIZE CONTEXT
# =========================================================

def normalize_context(
    attrs: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Builds the context used by the Word templates.

    The Word templates use the syntax:

        {{field_name}}

    The AGOL field names do not always match the names used
    inside the Word templates. We therefore preserve every
    AGOL field and add explicit template aliases.

    This is especially important for the Compliance Notice.
    """

    context: Dict[str, Any] = {}

    # ---------------------------------------------------------
    # Preserve every AGOL field exactly as received.
    # ---------------------------------------------------------

    for key, value in attrs.items():

        if value is None:
            context[key] = ""
        else:
            context[key] = value

    # ---------------------------------------------------------
    # Template aliases
    # ---------------------------------------------------------
    #
    # Compliance Notice template:
    #
    # {{name_premise}}
    # {{address__}}
    # {{Notice_No}}
    # {{owner_name}}
    # {{Surname}}
    # {{EditDate}}
    # {{Description_of_conduct}}
    # {{The_Section_or_Reg}}
    # {{Action_to_be_taken}}
    # {{period}}
    # {{EHP}}
    # {{signature}}
    #
    # Actual AGOL fields:
    #
    # premise_name
    # address
    # Name
    # Surname
    # EditDate
    # etc.
    # ---------------------------------------------------------

    context["name_premise"] = attrs.get(
        "premise_name"
    ) or ""

    context["address__"] = attrs.get(
        "address"
    ) or ""

    context["owner_name"] = attrs.get(
        "Name"
    ) or ""

    # These already match the AGOL field names.
    context["Notice_No"] = attrs.get(
        "Notice_No"
    ) or ""

    context["Surname"] = attrs.get(
        "Surname"
    ) or ""

    context["EditDate"] = attrs.get(
        "EditDate"
    ) or ""

    context["Description_of_conduct"] = attrs.get(
        "Description_of_conduct"
    ) or ""

    context["The_Section_or_Reg"] = attrs.get(
        "The_Section_or_Reg"
    ) or ""

    context["Action_to_be_taken"] = attrs.get(
        "Action_to_be_taken"
    ) or ""

    context["period"] = attrs.get(
        "period"
    ) or ""

    context["EHP"] = attrs.get(
        "EHP"
    ) or ""

    # ---------------------------------------------------------
    # Signature
    # ---------------------------------------------------------
    #
    # The Compliance Notice template calls this field
    # "signature".
    #
    # Prefer the actual recipient/person-in-charge signature
    # field if populated. Fall back to manager_signature and
    # then Ehp_signature.
    # ---------------------------------------------------------

    signature = (
        attrs.get("SIGNATURE_OF_PERSON_IN_CHARGE_")
        or attrs.get("manager_signature")
        or attrs.get("Ehp_signature")
        or ""
    )

    context["signature"] = signature

    # ---------------------------------------------------------
    # Useful aliases for common template naming differences.
    # ---------------------------------------------------------

    context["premise_name"] = attrs.get(
        "premise_name"
    ) or ""

    context["address"] = attrs.get(
        "address"
    ) or ""

    context["Name"] = attrs.get(
        "Name"
    ) or ""

    context["Surname"] = attrs.get(
        "Surname"
    ) or ""

    context["inspection_status"] = attrs.get(
        INSPECTION_STATUS_FIELD
    ) or ""

    context["compliance_status"] = attrs.get(
        COMPLIANCE_FIELD
    ) or ""

    context["action_taken"] = attrs.get(
        ACTION_TAKEN_FIELD
    ) or ""

    return context


# =========================================================
# SURVEY123 IMAGE ATTACHMENT HELPERS
# =========================================================

class ImagePlaceholder:
    """Temporary marker converted to docxtpl.InlineImage at render time."""

    def __init__(self, path: str, width_mm: float) -> None:
        self.path = path
        self.width_mm = width_mm


def normalize_attachment_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def attachment_matches_field(
    attachment: Dict[str, Any],
    field_name: str
) -> bool:
    field_token = normalize_attachment_name(field_name)
    if not field_token:
        return False

    for key in ("name", "keywords", "keyword", "fieldName", "field_name"):
        token = normalize_attachment_name(attachment.get(key, ""))
        if field_token in token:
            return True

    return False


def list_image_attachments(
    layer: FeatureLayer,
    object_id: Any
) -> List[Dict[str, Any]]:
    attachments = layer.attachments.get_list(oid=object_id) or []

    print(
        f"Found {len(attachments)} total attachment(s) "
        f"for OBJECTID {object_id}."
    )

    image_attachments = []

    for attachment in attachments:
        content_type = str(attachment.get("contentType", "")).lower()
        print(
            "  Attachment: "
            f"id={attachment.get('id')}, "
            f"name={attachment.get('name')!r}, "
            f"contentType={content_type!r}, "
            f"size={attachment.get('size')}"
        )

        if content_type.startswith("image/"):
            image_attachments.append(attachment)

    print(
        f"Found {len(image_attachments)} image attachment(s) "
        f"for OBJECTID {object_id}."
    )

    return image_attachments


def download_image_attachment(
    layer: FeatureLayer,
    object_id: Any,
    attachment: Dict[str, Any]
) -> str:
    attachment_id = attachment.get("id")
    if attachment_id is None:
        raise RuntimeError(f"Attachment has no id: {attachment}")

    image_dir = os.path.join(TEMP_DIR, "source_images", str(object_id))
    os.makedirs(image_dir, exist_ok=True)

    print(
        "Downloading image attachment: "
        f"OBJECTID={object_id}, attachment_id={attachment_id}, "
        f"name={attachment.get('name')!r}"
    )

    downloaded = layer.attachments.download(
        oid=object_id,
        attachment_id=attachment_id,
        save_path=image_dir
    )

    if isinstance(downloaded, str) and os.path.isfile(downloaded):
        local_path = downloaded
    else:
        expected_name = os.path.basename(str(attachment.get("name") or ""))
        candidate = os.path.join(image_dir, expected_name)
        if os.path.isfile(candidate):
            local_path = candidate
        else:
            files = [
                os.path.join(image_dir, name)
                for name in os.listdir(image_dir)
                if os.path.isfile(os.path.join(image_dir, name))
            ]
            if not files:
                raise RuntimeError(
                    "Attachment download returned no local image. "
                    f"API result: {downloaded!r}"
                )
            local_path = files[-1]

    if not os.path.exists(local_path):
        raise FileNotFoundError(f"Downloaded image does not exist: {local_path}")

    if os.path.getsize(local_path) == 0:
        raise RuntimeError(f"Downloaded image is empty: {local_path}")

    print(f"Downloaded image successfully: {local_path}")
    return local_path


def find_image_for_field(
    layer: FeatureLayer,
    object_id: Any,
    field_name: str,
    image_attachments: List[Dict[str, Any]]
) -> Optional[str]:
    matches = [
        attachment
        for attachment in image_attachments
        if attachment_matches_field(attachment, field_name)
    ]

    if not matches:
        print(
            f"No image attachment matched field '{field_name}' "
            f"for OBJECTID {object_id}."
        )
        return None

    if len(matches) > 1:
        print(
            f"WARNING: {len(matches)} image attachments matched "
            f"'{field_name}'. Using {matches[0].get('name')!r}."
        )

    return download_image_attachment(layer, object_id, matches[0])


def build_image_context(
    layer: FeatureLayer,
    object_id: Any,
    context: Dict[str, Any]
) -> Tuple[Dict[str, Any], List[str]]:
    """Download matching Survey123 image attachments."""

    image_attachments = list_image_attachments(layer, object_id)
    temporary_image_files: List[str] = []

    for field_name in IMAGE_FIELD_NAMES:
        local_path = find_image_for_field(
            layer,
            object_id,
            field_name,
            image_attachments
        )

        if not local_path:
            context[field_name] = ""
            if REQUIRE_SIGNATURE_IMAGES:
                raise RuntimeError(
                    f"Required image attachment for '{field_name}' "
                    f"was not found for OBJECTID {object_id}."
                )
            continue

        context[field_name] = ImagePlaceholder(
            local_path,
            SIGNATURE_IMAGE_WIDTH_MM
        )
        temporary_image_files.append(local_path)

    # Backward-compatible alias for templates still using {{signature}}.
    context["signature"] = ""
    for field_name in (
        "SIGNATURE_OF_PERSON_IN_CHARGE_",
        "manager_signature",
        "Ehp_signature",
    ):
        value = context.get(field_name)
        if isinstance(value, ImagePlaceholder):
            context["signature"] = value
            break

    return context, temporary_image_files


# =========================================================
# JINJA / DOCXTPL
# =========================================================

# Your Word templates use:
#
# ${{field}}
#
# Therefore the Jinja delimiters must be:
# variable_start_string="${{"
# variable_end_string="}}"


# =========================================================
# JINJA / DOCXTPL
# =========================================================
#
# Word templates must use docxtpl's native Jinja syntax:
#
#     {{field_name}}
#
# Do not use ${{field_name}} or ${field_name}.
# =========================================================

jinja_env = Environment(
    enable_async=False
)


def render_docx_template(
    template_path: str,
    context: Dict[str, Any],
    out_path: str
) -> str:

    print(f"Rendering template: {template_path}")

    template = DocxTemplate(template_path)
    render_context = dict(context)

    # Diagnose template variables before rendering.
    try:
        missing_variables = template.get_undeclared_template_variables(
            jinja_env=jinja_env,
            context=render_context,
        )
        if missing_variables:
            print(
                "WARNING: Template contains variables not present in the "
                f"context: {sorted(missing_variables)}"
            )
        else:
            print("All template variables are present in the context.")
    except Exception as e:
        print(f"Template-variable diagnostic could not be completed: {e}")

    for key, value in list(context.items()):
        if isinstance(value, ImagePlaceholder):
            render_context[key] = InlineImage(
                template,
                image_descriptor=value.path,
                width=Mm(value.width_mm)
            )

    template.render(render_context, jinja_env)
    template.save(out_path)
    return out_path


def print_template_context(
    context: Dict[str, Any],
    label: str
) -> None:
    """
    Prints the key values being supplied to a document.
    This makes it easy to diagnose template-field mismatches
    in Render logs.
    """

    print(
        "----------------------------------------"
    )

    print(
        f"Template context for {label}:"
    )

    important_keys = [
        "name_premise",
        "address__",
        "Notice_No",
        "owner_name",
        "Surname",
        "EditDate",
        "Description_of_conduct",
        "The_Section_or_Reg",
        "Action_to_be_taken",
        "period",
        "EHP",
        "signature",
    ]

    for key in important_keys:

        value = context.get(key)

        if isinstance(value, ImagePlaceholder):
            display_value = f"<IMAGE: {value.path}>"
        else:
            display_value = value

        print(
            f"  {key} = {display_value!r}"
        )

    print(
        "----------------------------------------"
    )


# =========================================================
# DOCUMENT BUSINESS RULES
# =========================================================

def choose_templates(
    attrs: Dict[str, Any]
) -> List[Tuple[str, str]]:
    """
    Determines which documents should be generated.

    RULE 1
    ---------------------------------------------------------
    inspection_status = completed
    AND
    compliance_status = compliant

    RULE 2
    ---------------------------------------------------------
    action_taken = issue_fine

    Generate:
        Fine


    RULE 3
    ---------------------------------------------------------
    action_taken = compliance_notice

    Generate:
        Compliance Notice


    RULE 4
    ---------------------------------------------------------
    action_taken = prohibtion_notice

    Generate:
        Prohibition Notice

    IMPORTANT:
    'prohibtion_notice' is intentionally spelled this way
    because this matches the current XLSForm value.
    """

    inspection_status = str(
        attrs.get(
            INSPECTION_STATUS_FIELD
        ) or ""
    ).strip().lower()


    compliance_status = str(
        attrs.get(
            COMPLIANCE_FIELD
        ) or ""
    ).strip().lower()


    action_taken = str(
        attrs.get(
            ACTION_TAKEN_FIELD
        ) or ""
    ).strip().lower()


    print(
        "----------------------------------------"
    )

    print(
        f"{INSPECTION_STATUS_FIELD} = "
        f"{inspection_status!r}"
    )

    print(
        f"{COMPLIANCE_FIELD} = "
        f"{compliance_status!r}"
    )

    print(
        f"{ACTION_TAKEN_FIELD} = "
        f"{action_taken!r}"
    )

    print(
        "----------------------------------------"
    )


    documents = []


    # =====================================================
    # COMPLETED + COMPLIANT
    # =====================================================

    if (
        inspection_status == "completed"
        and
        compliance_status == "compliant"
    ):


        documents.append(
                    (
                        REPORT_TEMPLATE_PATH,
                        "REPORT"
                    )
                )
        


    # =====================================================
    # FINE
    # =====================================================

    if action_taken == "issue_fine":

        documents.append(
            (
                FINE_TEMPLATE_PATH,
                "FINE"
            )
        )
        documents.append(
                            (
                                REPORT_TEMPLATE_PATH,
                                "REPORT"
                            )
                        )


    # =====================================================
    # COMPLIANCE NOTICE
    # =====================================================

    if action_taken == "compliance_notice":

        documents.append(
            (
                COMPLIANCE_NOTICE_TEMPLATE_PATH,
                "COMPLIANCE_NOTICE"
            )
        )
        documents.append(
                            (
                                REPORT_TEMPLATE_PATH,
                                "REPORT"
                            )
                        )


    # =====================================================
    # PROHIBITION NOTICE
    # =====================================================
    #
    # IMPORTANT:
    #
    # We deliberately use:
    #
    # prohibtion_notice
    #
    # because this is the value currently stored
    # by your XLSForm.
    # =====================================================

    if action_taken == "prohibtion_notice":

        documents.append(
            (
                PROHIBITION_NOTICE_TEMPLATE_PATH,
                "PROHIBITION_NOTICE"
            )
        )
        documents.append(
                            (
                                REPORT_TEMPLATE_PATH,
                                "REPORT"
                            )
                        )


    print(
        f"Documents selected: "
        f"{[label for _, label in documents]}"
    )

    return documents


# =========================================================
# ARC GIS CONNECTION
# =========================================================

def connect_gis() -> GIS:
    """
    Connects to ArcGIS Online.
    """

    require_env(
        "AGOL_USERNAME"
    )

    require_env(
        "AGOL_PASSWORD"
    )

    require_env(
        "FEATURE_LAYER_URL"
    )


    print(
        f"Connecting to ArcGIS Online: "
        f"{AGOL_URL}"
    )


    gis = GIS(
        AGOL_URL,
        AGOL_USERNAME,
        AGOL_PASSWORD
    )


    print(
        f"Logged into AGOL as: "
        f"{gis.users.me.username}"
    )


    return gis


# =========================================================
# GET FEATURE LAYER
# =========================================================

def get_feature_layer(
    gis: GIS
) -> FeatureLayer:
    """
    Opens the actual Feature Layer.

    FEATURE_LAYER_URL must end in /0,
    /1, /2, etc.
    """

    if not FEATURE_LAYER_URL:

        raise RuntimeError(
            "FEATURE_LAYER_URL is not configured."
        )


    print(
        "Opening Feature Layer URL:"
    )

    print(
        FEATURE_LAYER_URL
    )


    # -----------------------------------------------------
    # Make sure we're pointing to an individual layer.
    # -----------------------------------------------------

    final_part = (
        FEATURE_LAYER_URL
        .rstrip("/")
        .split("/")
        [-1]
    )


    if not final_part.isdigit():

        raise RuntimeError(
            "FEATURE_LAYER_URL must point to an "
            "individual Feature Layer.\n\n"
            "Expected something like:\n"
            "https://services6.arcgis.com/.../"
            "FeatureServer/0\n\n"
            "Current value:\n"
            f"{FEATURE_LAYER_URL}"
        )


    layer = FeatureLayer(
        FEATURE_LAYER_URL,
        gis=gis
    )


    # Force metadata loading.

    properties = layer.properties


    print(
        "Layer loaded successfully."
    )


    print(
        f"Object ID field: "
        f"{properties.get('objectIdField', 'UNKNOWN')}"
    )


    print(
        f"Has attachments: "
        f"{properties.get('hasAttachments', 'UNKNOWN')}"
    )


    return layer


# =========================================================
# VERIFY PROCESSING FIELDS
# =========================================================

def verify_processing_fields(
    layer: FeatureLayer
) -> None:

    fields = layer.properties.fields


    field_names = {
        field["name"]
        for field in fields
    }


    required_fields = [
        INSPECTION_STATUS_FIELD,
        COMPLIANCE_FIELD,
        ACTION_TAKEN_FIELD,
        REPORT_STATUS_FIELD,
    ]


    missing = [
        field
        for field in required_fields
        if field not in field_names
    ]


    if missing:

        raise RuntimeError(
            "The following required fields were not "
            f"found in the Feature Layer: {missing}\n\n"
            "Available fields:\n"
            f"{sorted(field_names)}"
        )


    print(
        "All required processing fields were found."
    )


# =========================================================
# VERIFY ATTACHMENTS
# =========================================================

def verify_attachments_enabled(
    layer: FeatureLayer
) -> None:

    try:

        has_attachments = (
            layer.properties.get(
                "hasAttachments"
            )
        )


        print(
            f"Feature Layer hasAttachments: "
            f"{has_attachments}"
        )


        if has_attachments is False:

            raise RuntimeError(
                "Attachments are NOT enabled on this "
                "Feature Layer."
            )


    except RuntimeError:

        raise


    except Exception as e:

        print(
            "Could not determine whether attachments "
            f"are enabled: {e}"
        )


# =========================================================
# OBJECT ID
# =========================================================

def get_object_id(
    attrs: Dict[str, Any],
    layer: FeatureLayer
) -> Any:

    object_id_field = (
        layer.properties.objectIdField
    )


    print(
        f"ArcGIS Object ID field: "
        f"{object_id_field}"
    )


    if object_id_field not in attrs:

        raise KeyError(
            f"Object ID field "
            f"'{object_id_field}' "
            "was not found in feature attributes."
        )


    return attrs[
        object_id_field
    ]


# =========================================================
# GET FEATURE BY OBJECT ID
# =========================================================

def get_feature_by_object_id(
    layer: FeatureLayer,
    object_id: Any
):
    """
    Retrieves one exact feature from AGOL.
    """

    object_id_field = (
        layer.properties.objectIdField
    )


    print(
        f"Retrieving exact feature: "
        f"{object_id_field} = {object_id}"
    )


    result = layer.query(
        where=(
            f"{object_id_field} = {object_id}"
        ),
        out_fields="*",
        return_geometry=False
    )


    if not result.features:

        raise RuntimeError(
            f"Could not find OBJECTID "
            f"{object_id} in the Feature Layer."
        )


    feature = result.features[0]


    print(
        f"Successfully retrieved OBJECTID "
        f"{object_id}."
    )


    return feature


# =========================================================
# ATTACHMENT CHECK
# =========================================================

def has_existing_attachments(
    layer: FeatureLayer,
    object_id: Any
) -> bool:

    try:

        attachments = (
            layer.attachments.get_list(
                oid=object_id
            )
        )


        exists = (
            attachments is not None
            and
            len(attachments) > 0
        )


        print(
            f"Existing attachments for "
            f"OBJECTID {object_id}: {exists}"
        )


        if attachments:

            print(
                f"Existing attachments: "
                f"{attachments}"
            )


        return exists


    except Exception as e:

        print(
            f"Could not check attachments "
            f"for OBJECTID {object_id}: {e}"
        )


        return False


# =========================================================
# ATTACH FILE
# =========================================================

def attach_file(
    layer: FeatureLayer,
    object_id: Any,
    file_path: str
) -> None:

    print(
        "----------------------------------------"
    )


    print(
        f"Attaching file to OBJECTID "
        f"{object_id}"
    )


    print(
        f"File: {file_path}"
    )


    # -----------------------------------------------------
    # Verify file exists
    # -----------------------------------------------------

    if not os.path.exists(
        file_path
    ):

        raise FileNotFoundError(
            f"Generated file does not exist: "
            f"{file_path}"
        )


    # -----------------------------------------------------
    # Verify file size
    # -----------------------------------------------------

    file_size = os.path.getsize(
        file_path
    )


    print(
        f"File size: {file_size} bytes"
    )


    if file_size == 0:

        raise RuntimeError(
            f"Generated file is empty: "
            f"{file_path}"
        )


    # -----------------------------------------------------
    # Upload attachment
    # -----------------------------------------------------

    try:

        print(
            f"Uploading attachment to "
            f"OBJECTID {object_id}..."
        )


        # IMPORTANT:
        #
        # Correct ArcGIS API for Python syntax:
        #
        # attachments.add(
        #     object_id,
        #     file_path
        # )
        #
        # DO NOT use:
        #
        # rel_objectid=object_id
        #
        result = layer.attachments.add(
            object_id,
            file_path
        )


        print(
            f"Attachment result: {result}"
        )


        # -------------------------------------------------
        # Check ArcGIS response
        # -------------------------------------------------

        if isinstance(
            result,
            dict
        ):

            add_result = result.get(
                "addAttachmentResult"
            )


            if add_result:

                success = add_result.get(
                    "success"
                )


                if success is not True:

                    raise RuntimeError(
                        "ArcGIS reported attachment "
                        f"upload failure: {result}"
                    )


        print(
            f"Attachment uploaded successfully "
            f"for OBJECTID {object_id}."
        )


        # -------------------------------------------------
        # Verify attachment exists
        # -------------------------------------------------

        attachments = (
            layer.attachments.get_list(
                oid=object_id
            )
        )


        filename = os.path.basename(
            file_path
        )


        matching = [
            attachment
            for attachment in attachments
            if attachment.get("name")
            == filename
        ]


        if matching:

            print(
                "Attachment verification successful."
            )


            print(
                f"Uploaded attachment: "
                f"{matching}"
            )


        else:

            print(
                "WARNING: ArcGIS returned a successful "
                "upload response, but the attachment "
                "could not be found during verification."
            )


            print(
                f"Current attachments: "
                f"{attachments}"
            )


    except Exception as e:

        print(
            "ATTACHMENT UPLOAD FAILED"
        )


        print(
            f"OBJECTID: {object_id}"
        )


        print(
            f"FILE: {file_path}"
        )


        print(
            f"ERROR: {e}"
        )


        print(
            traceback.format_exc()
        )


        raise


# =========================================================
# UPDATE REPORT STATUS
# =========================================================

def update_report_status(
    layer: FeatureLayer,
    feature,
    new_status: str
) -> None:

    print(
        f"Updating {REPORT_STATUS_FIELD} "
        f"to {new_status!r}"
    )


    feature.attributes[
        REPORT_STATUS_FIELD
    ] = new_status


    result = layer.edit_features(
        updates=[feature]
    )


    print(
        f"Status update result: {result}"
    )


    # -----------------------------------------------------
    # Validate update
    # -----------------------------------------------------

    if isinstance(
        result,
        dict
    ):

        update_results = (
            result.get(
                "updateResults"
            )
        )


        if update_results:

            first_result = (
                update_results[0]
            )


            if not first_result.get(
                "success",
                False
            ):

                raise RuntimeError(
                    "Failed to update report "
                    f"status: {result}"
                )


# =========================================================
# PROCESS ONE FEATURE
# =========================================================

def process_feature(
    layer: FeatureLayer,
    feature
) -> bool:
    """
    Processes exactly one Survey123 feature.
    """

    attrs = feature.attributes


    print(
        "\n"
        "========================================"
    )

    print(
        "PROCESSING NEW RECORD"
    )

    print(
        "========================================"
    )


    # -----------------------------------------------------
    # Print attributes
    # -----------------------------------------------------

    print(
        "Feature attributes:"
    )


    for key, value in attrs.items():

        print(
            f"  {key} = {value!r}"
        )


    # -----------------------------------------------------
    # Get Object ID
    # -----------------------------------------------------

    object_id = get_object_id(
        attrs,
        layer
    )


    print(
        f"OBJECTID: {object_id}"
    )


    # -----------------------------------------------------
    # Check current report status
    # -----------------------------------------------------

    current_status = str(
        attrs.get(
            REPORT_STATUS_FIELD
        ) or ""
    ).strip().lower()


    print(
        f"{REPORT_STATUS_FIELD} = "
        f"{current_status!r}"
    )


    if current_status == "generated":

        print(
            f"OBJECTID {object_id} is already "
            "marked as generated."
        )


        return False


    # -----------------------------------------------------
    # Determine required documents
    # -----------------------------------------------------

    documents = choose_templates(
        attrs
    )


    # -----------------------------------------------------
    # No matching rule
    # -----------------------------------------------------

    if not documents:

        print(
            f"No document rule matched "
            f"OBJECTID {object_id}."
        )


        print(
            "No documents will be generated."
        )


        return False


    # -----------------------------------------------------
    # Optional attachment check
    # -----------------------------------------------------

    if (
        SKIP_IF_ATTACHMENT_ALREADY_EXISTS
        and
        has_existing_attachments(
            layer,
            object_id
        )
    ):

        print(
            f"OBJECTID {object_id} "
            "already has attachments."
        )


        update_report_status(
            layer,
            feature,
            "generated"
        )


        return True


    # -----------------------------------------------------
    # Create DocxTemplate context
    # -----------------------------------------------------

    context = normalize_context(
        attrs
    )

    temporary_image_files: List[str] = []

    # -----------------------------------------------------
    # Generate documents
    # -----------------------------------------------------

    generated_files = []


    try:

        # -------------------------------------------------
        # Download Survey123 image/signature attachments.
        # -------------------------------------------------

        context, temporary_image_files = build_image_context(
            layer,
            object_id,
            context
        )

        print_template_context(
            context,
            "selected documents"
        )

        for template_path, label in documents:

            print(
                "\n"
                f"Generating {label} "
                f"for OBJECTID {object_id}"
            )


            # -------------------------------------------------
            # Create output path
            # -------------------------------------------------

            docx_path = out_docx_path(
                object_id,
                label=label
            )


            # -------------------------------------------------
            # Render DOCX
            # -------------------------------------------------

            render_docx_template(
                template_path,
                context,
                docx_path
            )


            print(
                f"Generated DOCX: "
                f"{docx_path}"
            )


            # -------------------------------------------------
            # Verify generated file
            # -------------------------------------------------

            if not os.path.exists(
                docx_path
            ):

                raise RuntimeError(
                    f"DOCX was not created: "
                    f"{docx_path}"
                )


            file_size = os.path.getsize(
                docx_path
            )


            print(
                f"DOCX size: "
                f"{file_size} bytes"
            )


            if file_size == 0:

                raise RuntimeError(
                    f"Generated DOCX is empty: "
                    f"{docx_path}"
                )


            # -------------------------------------------------
            # Attach DOCX
            # -------------------------------------------------

            attach_file(
                layer,
                object_id,
                docx_path
            )


            print(
                f"{label} attached successfully."
            )


            generated_files.append(
                docx_path
            )


        # -----------------------------------------------------
        # ONLY mark generated after ALL documents succeed
        # -----------------------------------------------------

        update_report_status(
            layer,
            feature,
            "generated"
        )


        print(
            f"OBJECTID {object_id} "
            "marked as generated."
        )


        return True


    finally:

        # -----------------------------------------------------
        # Delete temporary files
        # -----------------------------------------------------

        for docx_path in generated_files:

            try:

                if os.path.exists(
                    docx_path
                ):

                    os.remove(
                        docx_path
                    )


                    print(
                        f"Deleted temporary file: "
                        f"{docx_path}"
                    )


            except Exception as e:

                print(
                    f"Could not delete "
                    f"{docx_path}: {e}"
                )


        # -------------------------------------------------
        # Delete downloaded source images
        # -------------------------------------------------

        for image_path in temporary_image_files:

            try:
                if os.path.exists(image_path):
                    os.remove(image_path)
                    print(
                        f"Deleted temporary image: {image_path}"
                    )
            except Exception as e:
                print(
                    f"Could not delete image {image_path}: {e}"
                )


# =========================================================
# PROCESS FEATURE BY OBJECT ID
# =========================================================

def process_object_id(
    layer: FeatureLayer,
    object_id: Any
) -> bool:
    """
    Retrieves and processes one exact feature.
    """

    feature = get_feature_by_object_id(
        layer,
        object_id
    )


    return process_feature(
        layer,
        feature
    )


# =========================================================
# PROCESS UNPROCESSED RECORDS
# =========================================================

def build_unprocessed_where() -> str:

    return (
        f"({REPORT_STATUS_FIELD} IS NULL "
        f"OR {REPORT_STATUS_FIELD} = '')"
    )


def process_new_records(
    layer: FeatureLayer,
    limit: Optional[int] = None
) -> int:
    """
    Fallback/manual processing method.

    Searches for records where Coa_Notice is empty.
    """

    where_clause = (
        build_unprocessed_where()
    )


    print(
        "========================================"
    )


    print(
        "SEARCHING FOR UNPROCESSED RECORDS"
    )


    print(
        "========================================"
    )


    print(
        f"Querying Feature Layer with:"
    )


    print(
        where_clause
    )


    result = layer.query(
        where=where_clause,
        out_fields="*",
        return_geometry=False
    )


    features = (
        result.features
        if result
        and result.features
        else []
    )


    print(
        f"Found {len(features)} "
        "unprocessed record(s)."
    )


    processed_count = 0


    for feature in features:

        if (
            limit is not None
            and
            processed_count >= limit
        ):

            break


        try:

            successful = process_feature(
                layer,
                feature
            )


            if successful:

                processed_count += 1


        except Exception as e:

            print(
                "========================================"
            )


            print(
                "ERROR PROCESSING RECORD"
            )


            print(
                f"ERROR: {e}"
            )


            print(
                "========================================"
            )


            print(
                traceback.format_exc()
            )


    print(
        "\n"
        "========================================"
    )


    print(
        f"Completed. Successfully handled "
        f"{processed_count} record(s)."
    )


    print(
        "========================================"
    )


    return processed_count


# =========================================================
# COMMON AGOL PROCESSING SETUP
# =========================================================

def setup_processing() -> Tuple[GIS, FeatureLayer]:
    """
    Connects to AGOL and prepares the Feature Layer.
    """

    print(
        "\n"
        "========================================"
    )


    print(
        "REPORT PROCESSING STARTED"
    )


    print(
        "========================================"
    )


    # -----------------------------------------------------
    # Validate templates
    # -----------------------------------------------------

    validate_paths()


    # -----------------------------------------------------
    # Validate required environment variables
    # -----------------------------------------------------

    require_env(
        "FEATURE_LAYER_URL"
    )


    # -----------------------------------------------------
    # Connect to AGOL
    # -----------------------------------------------------

    gis = connect_gis()


    # -----------------------------------------------------
    # Get Feature Layer
    # -----------------------------------------------------

    layer = get_feature_layer(
        gis
    )


    # -----------------------------------------------------
    # Verify fields
    # -----------------------------------------------------

    verify_processing_fields(
        layer
    )


    # -----------------------------------------------------
    # Verify attachments
    # -----------------------------------------------------

    verify_attachments_enabled(
        layer
    )


    return gis, layer



# =========================================================
# WEBHOOK FIELD HELPERS
# =========================================================

def find_attribute_case_insensitive(
    attrs: Dict[str, Any],
    field_name: str
) -> Any:
    """
    Returns an attribute value without depending on the
    capitalization used by the Survey123 webhook payload.
    """

    wanted = str(field_name).lower()

    for key, value in attrs.items():

        if str(key).lower() == wanted:
            return value

    return None


def find_layer_field(
    layer: FeatureLayer,
    field_name: str
) -> Optional[str]:
    """
    Returns the actual field name from the Feature Layer,
    matching case-insensitively.
    """

    wanted = str(field_name).lower()

    for field in layer.properties.fields:

        actual_name = field.get("name")

        if (
            actual_name
            and
            str(actual_name).lower() == wanted
        ):
            return actual_name

    return None


def escape_sql_string(value: Any) -> str:
    """
    Escapes a string for an ArcGIS SQL where clause.
    """

    return str(value).replace("'", "''")


def get_global_id_from_webhook(
    payload: Dict[str, Any],
    webhook_feature: Dict[str, Any],
    webhook_attrs: Dict[str, Any]
) -> Optional[str]:
    """
    Survey123 can expose GlobalID in different locations
    depending on the webhook payload/configuration.

    Check all common locations.
    """

    candidates = []

    # 1. Normal attributes
    for key, value in webhook_attrs.items():

        if str(key).lower() == "globalid":

            candidates.append(value)

    # 2. Feature-level properties
    candidates.extend([
        webhook_feature.get("globalId"),
        webhook_feature.get("globalid"),
        webhook_feature.get("GlobalID"),
    ])

    # 3. Payload-level properties
    candidates.extend([
        payload.get("globalId"),
        payload.get("globalid"),
        payload.get("GlobalID"),
    ])

    for candidate in candidates:

        if candidate not in (None, ""):

            return str(candidate).strip()

    return None


def resolve_webhook_feature(
    layer: FeatureLayer,
    payload: Dict[str, Any]
):
    """
    Resolves the exact AGOL feature that triggered the
    Survey123 webhook.

    Resolution order:

    1. OBJECTID directly in webhook
    2. GlobalID in webhook -> AGOL query
    3. premise_id + Creator -> AGOL query
    4. premise_id alone -> AGOL query

    The fallback is necessary because Survey123 webhooks
    may omit OBJECTID and GlobalID from feature.attributes.
    """

    webhook_feature = payload.get("feature")

    if not isinstance(webhook_feature, dict):

        raise RuntimeError(
            "Survey123 webhook did not contain a valid "
            "'feature' object."
        )

    webhook_attrs = webhook_feature.get("attributes")

    if not isinstance(webhook_attrs, dict):

        raise RuntimeError(
            "Survey123 webhook feature did not contain "
            "a valid 'attributes' object."
        )

    print(
        "Survey123 webhook attribute keys:"
    )

    print(
        list(webhook_attrs.keys())
    )

    object_id_field = (
        layer.properties.objectIdField
    )

    # -----------------------------------------------------
    # Strategy 1: OBJECTID
    # -----------------------------------------------------

    object_id = find_attribute_case_insensitive(
        webhook_attrs,
        object_id_field
    )

    if object_id not in (None, ""):

        print(
            f"OBJECTID found directly in webhook: "
            f"{object_id}"
        )

        return get_feature_by_object_id(
            layer,
            object_id
        )

    print(
        "OBJECTID was not included in the "
        "Survey123 webhook attributes."
    )

    # -----------------------------------------------------
    # Strategy 2: GlobalID
    # -----------------------------------------------------

    global_id = get_global_id_from_webhook(
        payload,
        webhook_feature,
        webhook_attrs
    )

    if global_id:

        print(
            f"GlobalID found in webhook: "
            f"{global_id}"
        )

        global_id_field = None

        for field in layer.properties.fields:

            if (
                field.get("type")
                ==
                "esriFieldTypeGlobalID"
            ):

                global_id_field = field.get(
                    "name"
                )

                break

        if not global_id_field:

            global_id_field = find_layer_field(
                layer,
                "globalid"
            )

        if global_id_field:

            safe_global_id = escape_sql_string(
                global_id
            )

            where_clause = (
                f"{global_id_field} = "
                f"'{safe_global_id}'"
            )

            print(
                "Finding submitted feature using "
                "GlobalID."
            )

            print(
                f"GlobalID query: {where_clause}"
            )

            result = layer.query(
                where=where_clause,
                out_fields="*",
                return_geometry=False
            )

            if result.features:

                if len(result.features) > 1:

                    raise RuntimeError(
                        "More than one feature was returned "
                        f"for GlobalID {global_id}."
                    )

                feature = result.features[0]

                resolved_object_id = get_object_id(
                    feature.attributes,
                    layer
                )

                print(
                    f"OBJECTID found from GlobalID: "
                    f"{resolved_object_id}"
                )

                return feature

            print(
                "GlobalID was supplied but no AGOL "
                "feature matched it."
            )

    else:

        print(
            "GlobalID was not included in the "
            "Survey123 webhook payload."
        )

    # -----------------------------------------------------
    # Strategy 3: premise_id + Creator
    #
    # Your Survey123 webhook payload contains both of
    # these fields, and premise_id is generated uniquely
    # for the inspection.
    # -----------------------------------------------------

    premise_id = find_attribute_case_insensitive(
        webhook_attrs,
        "premise_id"
    )

    creator = find_attribute_case_insensitive(
        webhook_attrs,
        "Creator"
    )

    premise_id_field = find_layer_field(
        layer,
        "premise_id"
    )

    creator_field = find_layer_field(
        layer,
        "Creator"
    )

    if premise_id not in (None, "") and premise_id_field:

        safe_premise_id = escape_sql_string(
            premise_id
        )

        # Prefer premise_id + Creator when Creator is
        # available in both the webhook and layer.

        if (
            creator not in (None, "")
            and
            creator_field
        ):

            safe_creator = escape_sql_string(
                creator
            )

            where_clause = (
                f"{premise_id_field} = "
                f"'{safe_premise_id}'"
                f" AND "
                f"{creator_field} = "
                f"'{safe_creator}'"
            )

        else:

            where_clause = (
                f"{premise_id_field} = "
                f"'{safe_premise_id}'"
            )

        print(
            "Finding submitted feature using "
            "premise_id."
        )

        print(
            f"Fallback query: {where_clause}"
        )

        result = layer.query(
            where=where_clause,
            out_fields="*",
            return_geometry=False,
            order_by_fields="CreationDate DESC"
        )

        if result.features:

            if len(result.features) > 1:

                print(
                    f"Fallback query returned "
                    f"{len(result.features)} records. "
                    "Using the newest CreationDate record."
                )

            feature = result.features[0]

            resolved_object_id = get_object_id(
                feature.attributes,
                layer
            )

            print(
                f"OBJECTID found using premise_id: "
                f"{resolved_object_id}"
            )

            return feature

    # -----------------------------------------------------
    # Strategy 4: Creator + newest CreationDate
    #
    # This is a last-resort fallback only.
    # -----------------------------------------------------

    if (
        creator not in (None, "")
        and
        creator_field
    ):

        safe_creator = escape_sql_string(
            creator
        )

        where_clause = (
            f"{creator_field} = "
            f"'{safe_creator}'"
        )

        print(
            "WARNING: Falling back to the newest "
            "record created by the webhook creator."
        )

        print(
            f"Creator fallback query: "
            f"{where_clause}"
        )

        result = layer.query(
            where=where_clause,
            out_fields="*",
            return_geometry=False,
            order_by_fields="CreationDate DESC",
            result_record_count=1
        )

        if result.features:

            feature = result.features[0]

            resolved_object_id = get_object_id(
                feature.attributes,
                layer
            )

            print(
                f"OBJECTID found using Creator fallback: "
                f"{resolved_object_id}"
            )

            return feature

    # -----------------------------------------------------
    # Nothing worked
    # -----------------------------------------------------

    raise RuntimeError(
        "Could not identify the submitted Survey123 "
        "record.\n\n"
        f"OBJECTID field: {object_id_field}\n"
        f"Webhook keys: {list(webhook_attrs.keys())}\n"
        f"premise_id: {premise_id!r}\n"
        f"Creator: {creator!r}\n"
        "The webhook did not contain a usable OBJECTID "
        "or GlobalID and no reliable fallback record "
        "could be identified."
    )


# =========================================================
# SURVEY123 WEBHOOK
# =========================================================

@app.post("/webhook/survey123")
async def survey123_webhook(
    request: Request
):

    print(
        "\n========================================"
    )

    print(
        "SURVEY123 WEBHOOK RECEIVED"
    )

    print(
        "========================================"
    )

    try:

        # -------------------------------------------------
        # Read webhook body
        # -------------------------------------------------

        body = await request.body()

        print(
            f"Webhook body length: "
            f"{len(body)} bytes"
        )


        # -------------------------------------------------
        # Decode JSON
        # -------------------------------------------------

        import json

        try:

            payload = json.loads(
                body.decode("utf-8")
            )

        except Exception as e:

            raise RuntimeError(
                "Survey123 webhook body could not "
                f"be decoded as JSON: {e}"
            )


        print(
            f"Webhook eventType: "
            f"{payload.get('eventType')}"
        )


        # -------------------------------------------------
        # Prepare AGOL
        # -------------------------------------------------

        _, layer = setup_processing()


        # -------------------------------------------------
        # Resolve the exact submitted feature.
        #
        # IMPORTANT:
        # Survey123 is not guaranteed to include OBJECTID
        # in feature.attributes.
        #
        # resolve_webhook_feature() therefore tries:
        #
        # 1. OBJECTID
        # 2. GlobalID
        # 3. premise_id + Creator
        # 4. premise_id
        # 5. Creator/newest record
        # -------------------------------------------------

        feature = resolve_webhook_feature(
            layer,
            payload
        )


        # -------------------------------------------------
        # Get the real OBJECTID from AGOL
        # -------------------------------------------------

        object_id = get_object_id(
            feature.attributes,
            layer
        )


        print(
            f"Processing exact AGOL OBJECTID: "
            f"{object_id}"
        )


        # -------------------------------------------------
        # Process the exact feature
        # -------------------------------------------------

        successful = process_feature(
            layer,
            feature
        )


        print(
            "========================================"
        )

        print(
            "SURVEY123 WEBHOOK PROCESSING COMPLETE"
        )

        print(
            "========================================"
        )


        return {
            "success": True,
            "object_id": object_id,
            "processed": successful
        }


    except Exception as e:

        print(
            "\n========================================"
        )

        print(
            "SURVEY123 WEBHOOK PROCESSING FAILED"
        )

        print(
            "========================================"
        )

        print(
            traceback.format_exc()
        )


        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc()
            }
        )



# =========================================================
# MANUAL PROCESSING ENDPOINT
# =========================================================

@app.post("/process")
def process_reports():

    try:

        gis, layer = setup_processing()


        # -------------------------------------------------
        # Processing limit
        # -------------------------------------------------

        limit_env = os.getenv(
            "PROCESS_LIMIT"
        )


        limit = (
            int(limit_env)
            if limit_env
            else None
        )


        # -------------------------------------------------
        # Process unprocessed records
        # -------------------------------------------------

        processed = process_new_records(
            layer,
            limit=limit
        )


        return {
            "success": True,
            "processed_records": processed
        }


    except Exception as e:

        print(
            "\n"
            "========================================"
        )


        print(
            "MANUAL REPORT PROCESSING FAILED"
        )


        print(
            "========================================"
        )


        print(
            traceback.format_exc()
        )


        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": str(e),
                "traceback": traceback.format_exc()
            }
        )


# =========================================================
# HEALTH CHECK
# =========================================================

@app.get("/")
def health_check():

    return {
        "status": "online",
        "service": "Survey123 Report Generator",
        "webhook": "/webhook/survey123",
        "manual_processing": "/process"
    }
