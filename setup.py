from distutils.core import setup
setup(
  name = 'changeOffice',
  packages = ['changeOffice'],
  version = '0.1',
  description = 'change msOffice format from xls to xlsx, doc to docx, ppt to pptx',
    long_description=open('README.md').read(),
  author = 'Wang Qi',
  author_email = 'wangmarkqi@gmail.com',
  url = 'https://github.com/wangmarkqi/changeOffice',
  download_url = 'https://github.com/wangmarkqi/changeOffice/archive/0.1.tar.gz',
    install_requires=[ 'pywin32'],
  keywords = ['excel', 'word', 'ppt','xls to xlsx','doc to docx','ppt to pptx'],
  classifiers = [
    'Development Status :: 0.1',
    'Operating System :: Windows',
    'Intended Audience :: Accountant,Financial analyst Developers',
    'License :: MIT License',
    'Programming Language :: Python',
    'Programming Language :: Python :: Implementation',
    'Programming Language :: Python :: 3',
    'Programming Language :: Python :: 3.6',
    'Topic :: Data Process :: Libraries'
  ],
)