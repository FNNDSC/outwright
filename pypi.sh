#!/bin/bash

G_SYNOPSIS="

 NAME

	pypi.sh

 SYNOPSIS

	pypi.sh <ver>

 ARGS

	<ver>
	A version string to upload. Typically something like '0.20.22'.

 DESCRIPTION

	pypi.sh is a simple helper script to tag and upload a new version of pypi.sh


"

if (( $# != 1 )) ; then
    echo "$G_SYNOPSIS"
    exit 1
fi

VER=$1
DIR=$PWD
pandoc --from=markdown --to=rst --output=README.rst README.adoc
git commit -am "v${VER}"
git push origin main
git tag $VER
git push origin --tags

#rstcheck README.rst
python3 setup.py sdist
cd $DIR
twine upload dist/outlook_email_sender-${VER}.tar.gz


