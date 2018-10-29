# vue-json2excel

A lightweight vue js component uses [xlsx](https://github.com/SheetJS/js-xlsx) library to export json data into an excel file.


## Project setup
```
npm install vue-json2-excel --save
```
or
```
yarn add vue-json2-excel
```
## Usage
```
    <template>
        <json2-excel
        :data="data"
        :header="headers"
        :details="{text:'This is a test text appears in details section of execl file'}">
            export
        </json2-excel>
    </template>
    
    <script>
        import Json2Excel from 'vue-json2excel';
        
        export default{
            data(){
                return{
                    data:[['aref','hosseinikia'],['feri','tajedin']],
                    headers:['first_name','last_name'],
                }
            },
            components:{
                Json2Excel
            }
        }
    </script> 
```