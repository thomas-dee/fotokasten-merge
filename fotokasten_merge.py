import os
import json
import argparse
import uuid
import shutil
import tempfile
import zipfile
import hashlib
import time
from PIL import Image
from pillow_heif import register_heif_opener

class FotokastenMerger():
    def __init__(self, file_name):
        self.merge_config = None
        self.source_projects = {}
        self.month_mapping = {
            "Cover": 0,
            "January": 1,
            "February": 2,
            "March": 3,
            "April": 4,
            "May": 5,
            "June": 6,
            "July": 7,
            "August": 8,
            "September": 9,
            "October": 10,
            "November": 11,
            "December": 12
        }

        register_heif_opener()

        with open(file_name, "r") as json_f:
            self.merge_config = json.load(json_f)

            self.temp_dir = tempfile.TemporaryDirectory()
            self.metadata = None

            #for key, id in self.merge_config["source-mapping"].items():
            #    self.__read_source_project(id)

    def __del__(self):
        self.temp_dir.cleanup()

    def __read_source_project(self, prj):
        if self.metadata is None:
            with open(os.path.join(self.temp_dir.name, prj, "META-INF", "metadata.json"), "r", encoding="utf-8-sig") as metadata_f:
                self.metadata = json.load(metadata_f)

        with open(os.path.join(self.temp_dir.name, prj, "PROJECT", "projectDescriptor.json"), "r", encoding="utf-8-sig") as project_f:
            self.source_projects[prj] = json.load(project_f)

    def Unpack(self):
        for prj in self.merge_config["source-mapping"].values():
            with zipfile.ZipFile("./%s" % prj, 'r') as zip_ref:
                print("Unpacking %s..." % prj)
                prj_path = os.path.join(self.temp_dir.name, prj)
                os.mkdir(prj_path)
                zip_ref.extractall(prj_path)
                self.__read_source_project(prj)

    def __deep_copy_page(self, page_cfg):
        new_page_cfg = json.loads(json.dumps(page_cfg))
        new_page_cfg["id"] = str(uuid.uuid4())

        for layer in new_page_cfg["layers"]:
            for element in layer["elements"]:
                element["id"] = str(uuid.uuid4())

        return new_page_cfg

    def Merge(self):
        # compare headers of source projects
        new_project_skeleton = None

        for cfg in self.source_projects.values():
            if new_project_skeleton is None:
                new_project_skeleton = {}
                new_project_skeleton["descriptor"] = cfg["descriptor"]
                new_project_skeleton["descriptor"]["id"] = str(uuid.uuid4())
                new_project_skeleton["descriptor"]["name"] = self.merge_config["name"]
                new_project_skeleton["productId"] = cfg["productId"]
                new_project_skeleton["options"] = cfg["options"]
                new_project_skeleton["compositionMetaInfo"] = cfg["compositionMetaInfo"]
                new_project_skeleton["pages"] = []

        new_project_skeleton["pages"] = [{}] * 13

        for month, cfg in self.merge_config["pages"].items():
            dst_page = self.month_mapping[month]
            for project_short, src_page_name in cfg.items():
                project_key = self.merge_config["source-mapping"][project_short]
                src_page = self.month_mapping[src_page_name]
                new_project_skeleton["pages"][dst_page] = self.__deep_copy_page(self.source_projects[project_key]["pages"][src_page])
                new_project_skeleton["pages"][dst_page]["sourceProject"] = project_key

        return new_project_skeleton

    def __md5(self, fname):
        hash_md5 = hashlib.md5()
        with open(fname, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def __store_as_jpeg(self, filepath):
        try:
            image=Image.open(filepath)
            image.save(filepath + ".jpg", "JPEG")
        except:
            pass

        return filepath + ".jpg"

    def __fix_rotation(self, filepath):
        try:
            image=Image.open(filepath)
            orientation = 274 # exif tag
            
            exif = image._getexif()
            if exif is not None:
                need_save = False
                if exif[orientation] == 3:
                    image=image.rotate(180, expand=True)
                    need_save = True
                elif exif[orientation] == 6:
                    image=image.rotate(270, expand=True)
                    need_save = True
                elif exif[orientation] == 8:
                    image=image.rotaste(90, expand=True)
                    need_save = True

                if need_save:
                    image_format = image.format if image.format is not None else "JPEG"
                    image.save(filepath, image_format)
                    image.close()

        except:
            # cases: image don't have getexif
            pass


    def WriteNewProjectPrj(self, project_cfg):
        # create sceleton dir
        new_prj_filename = os.path.join(".", project_cfg["descriptor"]["name"] + ".prj")
        print("Creating \"%s\"..." % new_prj_filename)

        with zipfile.ZipFile(new_prj_filename, "w", zipfile.ZIP_STORED) as zip_f:
            zip_f.writestr("mimetype", "application/pgx-project+zip".encode("utf-8-sig"), zipfile.ZIP_STORED)
            zip_f.writestr("META-INF\\file-version", "1.0.0.0".encode("utf-8-sig"), zipfile.ZIP_DEFLATED)
            zip_f.writestr("META-INF\\metadata.json", json.dumps(self.metadata, indent=4).encode("utf-8-sig"), zipfile.ZIP_DEFLATED)

            # images
            written_images = []
            page_idx = 0
            for page_cfg in project_cfg["pages"]:
                page_idx = page_idx + 1
                page_image_path = os.path.join(self.temp_dir.name, page_cfg["sourceProject"], "PROJECT", "IMAGES")
                #page_cfg.pop("sourceProject")

                for layer in page_cfg["layers"]:
                    for element in layer["elements"]:
                        if element["type"] == "PICTURE":
                            if "picture" in element and element["picture"] is not None:
                                picture_id = element["picture"]["id"]
                                picture_name = element["picture"]["name"]

                                picture_path = os.path.join(page_image_path, picture_id)

                                if picture_name.endswith(".heic") or picture_name.endswith(".HEIC"):
                                    element["picture"]["mimeType"] = "image/jpeg"
                                    picture_path = self.__store_as_jpeg(picture_path)

                                if picture_id.startswith("PGFileSystemSourcePrefix+"):
                                    # MacOS-Fotokasten handles rotations differently
                                    self.__fix_rotation(picture_path)

                                with open(picture_path, "rb") as pic_f:
                                    pic_data = pic_f.read()
                                    element["picture"]["id"] = hashlib.sha1(pic_data).hexdigest().upper()

                                    if not element["picture"]["id"] in written_images:
                                        pic_info = zipfile.ZipInfo("PROJECT\\IMAGES\\" + element["picture"]["id"])
                                        pic_info.comment = picture_name.encode("utf-8")
                                        pic_info.compress_type = zipfile.ZIP_STORED

                                        pic_time = element["picture"]["exifDate"] if "exifDate" in element["picture"] and element["picture"]["exifDate"] is not None else element["picture"]["lastModified"]
                                        element["picture"]["lastModified"] = pic_time

                                        pic_info.date_time = time.gmtime(pic_time / 1000)[:6]
                                        zip_f.writestr(pic_info, pic_data)

                                        written_images.append(element["picture"]["id"])

                        elif element["type"] == "TEXT" and page_idx > 1:
                            # allow to remove Text if it WAS a cover page
                            element["permissions"]["explicitPermissions"] = {}

            # project file
            zip_f.writestr("PROJECT\\projectDescriptor.json", json.dumps(project_cfg).encode("utf-8-sig"), zipfile.ZIP_DEFLATED)

        
    def WriteNewProject(self, project_cfg):
        # create sceleton dir
        with tempfile.TemporaryDirectory() as prj_path:
            print("Creating prj sceleton...")
            os.mkdir(os.path.join(prj_path, "META-INF"))
            os.mkdir(os.path.join(prj_path, "PROJECT"))

            images_path = os.path.join(prj_path, "PROJECT", "IMAGES")
            os.mkdir(images_path)

            with open(os.path.join(prj_path, "mimetype"), "w", encoding="utf-8-sig") as mimetype_f:
                mimetype_f.write("application/pgx-project+zip")

            with open(os.path.join(prj_path, "META-INF", "file-version"), "w", encoding="utf-8-sig") as mimetype_f:
                mimetype_f.write("1.0.0.0")

            with open(os.path.join(prj_path, "META-INF", "metadata.json"), "w", encoding="utf-8-sig") as metadata_f:
                json.dump(self.metadata, metadata_f, indent=4)

            # images
            print("Copying images...")
            for page_cfg in project_cfg["pages"]:
                page_image_path = os.path.join(self.temp_dir.name, page_cfg["sourceProject"], "PROJECT", "IMAGES")
                for layer in page_cfg["layers"]:
                    for element in layer["elements"]:
                        if element["type"] == "PICTURE" and "picture" in element and element["picture"] is not None:
                            picture_id = element["picture"]["id"]
                            picture_path = os.path.join(page_image_path, picture_id)
                            element["picture"]["id"] = self.__md5(picture_path).upper()
                            new_picture_path = os.path.join(images_path, element["picture"]["id"])
                            if not os.path.exists(new_picture_path):
                                shutil.copy(picture_path, new_picture_path)

            # project file
            print("Writing project file...")
            with open(os.path.join(prj_path, "PROJECT", "projectDescriptor.json"), "w", encoding="utf-8-sig") as project_f:
                json.dump(project_cfg, project_f)

            # zip
            new_prj_filename = os.path.join(".", project_cfg["descriptor"]["name"])
            print("Creating \"%s.prj\"..." % new_prj_filename)
            shutil.make_archive(os.path.join(".", project_cfg["descriptor"]["name"]), 'zip', prj_path)
            shutil.move(new_prj_filename + ".zip", new_prj_filename + ".prj")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", help="Config file", default="./merge-config.json")

    args = parser.parse_args()

    if (args.config):
        fk_merger = FotokastenMerger(args.config)
        fk_merger.Unpack()
        fk_merger.WriteNewProjectPrj(fk_merger.Merge())

