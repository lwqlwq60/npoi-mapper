#!/bin/bash

echo 'start to delete old packages...'

rm -rf ./nupkgs

echo 'old packages are deleted.'

echo 'start to pack nuget packages...'

dotnet pack -c release --output nupkgs

echo 'package built.'

# nuget push ./nupkgs/*.nupkg -k somekey -s https://api.nuget.org/v3/index.json

# dotnet nuget push ./nupkgs/*.nupkg -k 50301748-C7E9-11EA-AC68-C9DCFB2CA371 -s http://nuget.garrettmotion.io/v3/index.json

# echo 'package pushed.'
