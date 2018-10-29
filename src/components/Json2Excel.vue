<template>
    <div @click="exportToExcel" class="main">
        <slot></slot>
    </div>
</template>

<script>
    import xlsx from 'xlsx';
    export default {
        props: {
            header: {
                type: Array,
                default() {
                    return [];
                }
            },
            data: {
                type: Array,
                default() {
                    return [];
                }
            },
            details: {
                type: Object,
                default() {
                    return {};
                }
            }
        },
        name: "Json2Excel",
        data() {
            return {
                workBook: {},
                workSheet: {}
            }
        },
        methods: {
            init() {
                this.workBook = xlsx.utils.book_new();
            },
            makeDetails() {
                let width = this.header.length, height = Object.keys(this.details).length;

                this.workSheet['!merges'] = [];
                this.workSheet['!merges'].push({s: {r: 0, c: 0}, e: {r: height - 1, c: width - 1}});

                let header = xlsx.utils.encode_cell({r: 0, c: 0});

                let value = '___';
                for (let k in this.details)
                    if (this.details.hasOwnProperty(k))
                        value += k + '  :  ' + this.details[k] + '___\n';

                this.workSheet[header] = {v: value};
            },
            makeHeader() {
                this.header.forEach((h, index) => {
                    let cellRef = xlsx.utils.encode_cell({
                        r: Object.keys(this.details).length,
                        c: index
                    });
                    this.workSheet[cellRef] = {v: h};
                })
            },
            makeData() {
                this.data.forEach((row, rowIndex) => {
                    row.forEach((col, colIndex) => {
                        let cellRef = xlsx.utils.encode_cell({
                            r: rowIndex + Object.keys(this.details).length + 1,
                            c: colIndex
                        });
                        this.workSheet[cellRef] = {v: col}
                    })
                })
            },
            makeRange() {
                let rows = Object.keys(this.details).length + this.data.length,
                    cols = this.header.length;

                let range = {
                    s: {r: 0, c: 0},
                    e: {r: rows, c: cols}
                };

                this.workSheet['!ref'] = xlsx.utils.encode_range(range);

            },
            makeExport() {
                this.workBook.SheetNames.push('test');
                this.workBook.Sheets.test = this.workSheet;

                xlsx.writeFile(this.workBook, 'test.xlsx');
            },
            exportToExcel() {
                this.makeDetails();
                this.makeHeader();
                this.makeData();
                this.makeRange();
                this.makeExport();
            }
        },
        created() {
            this.init();
        },
    }
</script>

<style scoped>
    .main:hover{
        cursor: pointer;
    }
</style>