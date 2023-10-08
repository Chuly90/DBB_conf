# DBB_conf
This is the repository for the paper with title: **A pragmatic methodology to evaluate the configuration for a double busbar substation in an electrical grid** by Jorge Mola & Delio Gomez.

*Abstract— 
This paper addresses the optimization of double busbar substations with multiple fields to prevent overcurrents through the coupler and therefore enhance grid reliability. A matrix-based method is proposed to estimate current coupling, and real-life simulations are conducted on the Colombian bulk power system. The results demonstrate the effectiveness of the approach in identifying optimal configurations, ensuring secure and flexible operation. The study provides valuable insights into enhancing power grid stability and efficiency.*

Keywords— data, fast calculation, historian, matrix, pragmatic, Python, substation

Link to IEEEXplore: [soon](https://ieeexplore.ieee.org/)

**Content:**
* **Example** (folder): Historical data (P and Q) for an eight bays substation. The code in this repository will only run with this information. If you are interested in setting this code for your system, you must change the function `ConsultarDatos()` such that it is possible to get the historical data of P and Q flow through the bays of your substation of choice.

* **Config_SEs_DobleBarra.pfd** (PowerFactory ComPython): It is the PowerFactory user interface to run this code. The version in this repository doesn't query historical data and only works with the files in the Example folder, therefore, even if you run this ComPython using PowerFactory it won't do anything with it and, instead, will look for the folder Example to get the example files. Again, if you are interested in setting this code for your system, you must change the function `ConsultarDatos()` such that it is possible to get the historical data of P and Q flow through the bays of your substation of choice using PowerFactory as user interface.
  
* **Double_Busbar_SE_Conf.py** (Python file): Python script where the logic of this methodology is programmed.

* **requirements.txt** (text file): Text file with the required Python libraries.
