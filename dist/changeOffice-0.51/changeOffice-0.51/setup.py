from setuptools import setup
setup(
  name = 'changeOffice',
  packages = ['changeOffice'],
  version = '0.51',
  description = 'change msOffice format from xls to xlsx, doc to docx, ppt to pptx',
  author = 'Wang Qi',
  author_email = 'wangmarkqi@gmail.com',
  url = 'https://github.com/wangmarkqi/changeOffice',
  download_url = 'https://github.com/wangmarkqi/changeOffice.git',
    install_requires=[ 'pywin32>=214'],
  keywords = ['excel', 'word', 'ppt','xls to xlsx','doc to docx','ppt to pptx'],
  classifiers = [],
)

'''
# 上传source 包
python setup.py sdist build
python setup.py sdist upload
# 上传pre-compiled包
python setup.py bdist_wheel --universal
python setup.py bdist_wheel upload
'''