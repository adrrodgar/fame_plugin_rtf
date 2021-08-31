try:
    from oletools import rtfobj
    HAVE_OLETOOLS = True
except ImportError:
    HAVE_OLETOOLS = False

from fame.common.exceptions import ModuleInitializationError
from fame.core.module import ProcessingModule
import hashlib

class RTF_plugin(ProcessingModule):

    acts_on = ["rtf"]
    name = "RTF reversing"
    description = "Get objects from RTF files"
    config = []

    def initialize(self):
        if not HAVE_OLETOOLS:
            raise ModuleInitializationError(self, 'Missing dependency: oletools')
        return True

    def each_with_type(self, target, file_type):

        self.results = dict()
        content = open(target,'rb').read()
        rtfdata = content.decode('UTF-8')
        replace_content = rtfdata.replace('datastore', 'objdata')
        open(target, 'wb').write(str.encode(replace_content))
        self.results['objects'] = list()

        for index, orig_len, data in rtfobj.rtf_iter_objects(target):
            object_dict = dict()
            object_dict['sha256'] = hashlib.sha256(data).hexdigest()
            object_dict['code'] = data.decode(errors='replace')
            self.results['objects'].append(object_dict)

        if len(self.results['objects']) > 0:
            return True
        else:
            return False
