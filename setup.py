from setuptools import setup, find_packages

setup(
    name="strategic-console",
    version="1.0.0",
    python_requires='>=3.10.0,<3.11.0',  # Force Python 3.10.x
    install_requires=[
        'Flask==3.0.0',
        'pandas>=2.2.0',
        'plotly==5.18.0',
        'openpyxl==3.1.2',
        'gunicorn==21.2.0',
        'python-dotenv==1.0.0',
        'google-api-python-client==2.108.0',
        'google-auth-oauthlib==1.1.0',
        'google-auth-httplib2==0.1.1'
    ],
    packages=find_packages(),
)
