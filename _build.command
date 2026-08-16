#!/bin/bash

filepath=$(cd "$(dirname "$0")"; pwd)
cd "$filepath"

source env-setup
pmc_finit
pmc_itext
pmc_isys
pmc_build

cp pokecrystal11_debug.sav build/pokecrystal11_debug.sav