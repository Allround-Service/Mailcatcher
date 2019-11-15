try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup


setup(
    name='mailcatcher',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    py_modules=[
        'mailcatcher',
    ],
)
