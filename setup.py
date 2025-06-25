from setuptools import setup

setup(
    name="createur-dossier-expertise",
    version="1.0",
    description="Crée automatiquement une structure de dossiers pour les expertises",
    author="Sagenco",
    author_email="info@sagette.ch",
    py_modules=["creer_dossier_expertise"],
    entry_points={
        'console_scripts': [
            'creer-dossier-expertise=creer_dossier_expertise:main',
        ],
    },
)
