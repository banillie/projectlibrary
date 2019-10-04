from setuptools import setup, find_packages

setup(name="projectlibrary",
      version="0.1",
      description="Test Project Description",
      author="Mr Grant",
      install_requires=[
          'datamaps @ https://github.com/hammerheadlemon/datamaps/archive/master.zip#egg=datamaps'
      ]
      packages=find_packages()
      )
