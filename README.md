# WordGen
介绍： ListenAI 旗下 chipsky 的专用文档生成器，根据IC生成的XLS文件，生成模板，再由模板渲染成最终文档.局限性： XLS文件中关键字发生改变的时候，会导致模板发生变化，从而无法对原有模板进行渲染. 作者： EASON WANG



Release command:

​				pyinstaller -F main.py -n WordGen -i template.ico

Code:

​				https://github.com/MrainFly/WordGen.git