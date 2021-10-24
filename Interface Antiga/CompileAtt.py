import os
from shutil import copyfile, copytree

if not os.path.isdir("dist/"):
    print("Directory does not exist!")
else:
    copyfile("Equip_params.txt", "dist\\Equip_params.txt")
    copyfile("Temp_params.txt", "dist\\Temp_params.txt")
    copyfile("Temp_params.txt", "dist\\Temp_params_low.txt")
    copyfile("Filetemp_params.txt", "dist\\Filetemp_params.txt")
    copyfile("Icone.png", "dist\\Icone.png")
    copytree("Applications", "dist\\Applications")