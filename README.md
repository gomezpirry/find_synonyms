# find_synonyms

Find synonyms in nodes of an obo file based on id list. Program written in python 3

## Usage

### Run locally 

Install dependencies (Python 3)

```
pip install -r requeriments.txt
```

Find synonyms in __input_file.obo__ reading the ids of __input_file.xls__. The input file xls should be have a cell with "ID" text, the cells below this will be taken as ids
__Note:__ Run in this way, you rewrite original __input_file.xls__

```
cd find_synonyms
python3 find_synonyms -i input_file.obo -x input_file.xls
```

Find synonyms in __input_file.obo__ reading the ids of __input_file.xls__. The results are stored in __output_file.xls__. The input file xls should be have a cell with "ID" text, the cells below this will be taken as ids

```
cd find_synonyms
python3 find_synonyms -i input_file.obo -x input_file.xls -o output_file.xls
```

### Run in Docker

Build dockerfile 

```
cd find_synonyms
docker build -t find_synonyms .
```

Run the docker repository. 

The option __-v__ allows connect a host folder with a docker volume. 
As __-v__ option argument pass the host folder path where the obo file is
followed by(:) and the volume path in docker (you decide the name of 
volume path). With this option, the host folder and volume docker are 
synchronized (shared files). The value of input obo __-i__, input xls __x__ and output __o__ arguments are the volume path followed by input obo file name, input xls filename and output file name respectively
   
```
docker run -v /path/to/folder:/path_to_docker find_synonyms -i /path_to_docker/input_file.obo -x /path_to_docker/input_file.xls -o /path_to_docker/output_file.xls
```
