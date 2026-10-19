import json
import inflection


class JsonConverter(object):

    def dump(self, element):
        """Dumps an object to a json formatted string,
        in the process all snake_case keys are converted to camelCase.

        :param element: object
        :type element: object
        :return: json formatted string, representing the the object element
        """

        json_string = json.dumps(element, default=self.__jsonDefault)

        # convert snake_case to camelCase

        # load jsonString into jsonObject, while doing so, convert all keys to camelCase
        json_object = json.loads(json_string, object_hook=self.__convert_keys_to_camelCase)
        # dump back to jsonString
        json_string = json.dumps(json_object)

        return json_string

    def load(self, element):
        """Loads an object given a json formatted string,
        in the process all camelCase keys are converted to snake_case.

        :param element: json formatted string
        :type element: str
        :return: object, representing json formatted string
        """

        return json.loads(element, object_hook=self.__convert_keys_to_snake_case)

    def __jsonDefault(self, obj):
        """Aids to dump custom objects to json, by converting them to a dictionary,
        removing all fields which are set to None in the process."""
        return dict(filter(lambda item: item[1] is not None, obj.__dict__.items()))

    def __convert_keys_to_camelCase(self, obj):
        """Converts all keys to camelCase."""
        # to prevent iterating over keys twice,
        # as they are newly added, they are also added to the iterator
        number_of_keys = 0
        for key in obj.keys():
            number_of_keys += 1
        current = 1
        new_obj = dict()
        # replace keys
        for key in obj.keys():
            if current > number_of_keys:
                break
            current += 1
            # camelize and remove leading '_'
            new_key = inflection.camelize(key)[1:]
            # if new_key != key:
            new_obj[new_key] = obj[key]
                # del obj[key]
        return new_obj

    def __convert_keys_to_snake_case(self, obj):
        """Converts all keys to snake_case."""
        # to prevent iterating over keys twice,
        # as they are newly added, they are also added to the iterator
        number_of_keys = 0
        for key in obj.keys():
            number_of_keys += 1
        current = 1
        new_obj = dict()
        # replace keys
        for key in obj.keys():
            if current > number_of_keys:
                break
            current += 1
            # switch to snake_case
            new_key = inflection.underscore(key)
            # if new_key != key:
            new_obj[new_key] = obj[key]
                # del obj[key]
        return new_obj
