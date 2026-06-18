import cloudinary.uploader
from cloud.config import *

def subir_pdf(ruta_pdf,nombre_archivo):

    resultado = cloudinary.uploader.upload(
        ruta_pdf,
        resource_type="raw",
        format="pdf",
        public_id=nombre_archivo,
        overwrite=True
    )

    url = resultado["secure_url"]

    url = url.replace(
        "/raw/upload/",
        "/raw/upload/fl_attachment:false/"
    )

    return {
        "url":url,
        "cloud_id":resultado["public_id"]
    }


def eliminar_pdf(cloud_id):

    cloudinary.uploader.destroy(
        cloud_id,
        resource_type="raw"
    )