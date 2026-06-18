import cloudinary.uploader
from cloud.config import *

def subir_imagen(ruta_imagen, nombre_archivo):

    resultado = cloudinary.uploader.upload(
        ruta_imagen,
        resource_type="image",
        public_id=nombre_archivo,
        overwrite=True
    )

    return {
        "url": resultado["secure_url"],
        "cloud_id": resultado["public_id"]
    }

def eliminar_imagen(cloud_id):

    cloudinary.uploader.destroy(
        cloud_id,
        resource_type="image"
    )