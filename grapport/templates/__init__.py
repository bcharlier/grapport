# import all the py file in the grapport/templates folder

import os

__all__ = [os.path.basename(os.path.normpath(x[0])) for x in os.walk(os.path.dirname(__file__)) if x[0] != os.path.dirname(__file__) and "__pycache__" not in x[0]]
