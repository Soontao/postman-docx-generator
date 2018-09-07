import json
import postman_schema


def parsePostmanJson(path):
  result = ""
  with open(path, "r") as f:
    json_string = f.read()
    result = postman_schema.postman_object_from_dict(json.loads(json_string))
  return result
