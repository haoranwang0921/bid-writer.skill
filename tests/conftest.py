import os
import sys

# 使 tests 可直接 import scripts 下的模块
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts"))
