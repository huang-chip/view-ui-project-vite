<template>
  <div>
    {{ fields }}
    <Form :model="fields" :label-width="150" :rules="ruleCustom">
      <FormItem label="记录年份：" prop="recordYear">
      <Input v-model="fields.header[0].recordYear" placeholder="输入记录年份..."></Input>
    </FormItem>
    <FormItem label="记录编号：" prop="recordNo">
      <Input v-model="fields.header[0].recordNo" placeholder="输入记录编号..."></Input>
    </FormItem>
    <FormItem label="当前页码：" prop="currPage">
      <Input v-model="fields.header[0].currPage" placeholder="输入当前页码..."></Input>
    </FormItem>
    <FormItem label="总页数：" prop="totalPage">
      <Input v-model="fields.header[0].totalPage" placeholder="输入总页数..."></Input>
    </FormItem>
    <FormItem label="公司：" prop="company">
      <Input v-model="fields.basicInfo[0].company" placeholder="输入公司名称..."></Input>
    </FormItem>
    <FormItem label="温度：" prop="temperature">
      <Input v-model="fields.basicInfo[0].temperature" placeholder="输入温度..."></Input>
    </FormItem>
    <FormItem label="湿度：" prop="humidity">
      <Input v-model="fields.basicInfo[0].humidity" placeholder="输入湿度..."></Input>
    </FormItem>
<!-- 可以添加的 -->

      </Form>
    <button @click="generate">生成模版A</button>
  </div>
</template>

<script>
import docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import JSZipUtils from "jszip-utils";
import { saveAs } from "file-saver";
export default {
  name: "CreateMoudleA",
  components: {},

  props: {},

  data() {
    return {
      fields:{
      header: [
        {
          recordYear: "2024",
          recordNo: "012",
          currPage: "1",
          totalPage: "5",
        },
      ],
      basicInfo: [
        {
          company: "宁夏北控睿源再生资源有限公司",
          temperature: "9.4",
          humidity: "20.2",
        },
      ],
      tableInfo: [
        {
          // "index": 'a',
          sampleName: "LD-200S固定式气体探测器",
          factory: "南京更佳电子科技有限公司",
          targetGas: "可燃气",
          scaleRange: "0-100％LEL"
        },
        {
          // "index": 'b',
          sampleName: "LD-200S固定式气体探测器",
          factory: "南京更佳电子科技有限公司",
          targetGas: "H2S",
          scaleRange: "0-100ppm"
        },
      ],
      },
      tableObj: {
        sampleName:'',
        factory: '',
        targetGas: '',
        scaleRange: '',
      }
    };
  },
  computed: {},
  watch: {},

  mounted() {},
  methods: {
    addIndex18(){
      let that = this;
      const indexes = Array.from('abcdefghijklmnopqr');
      // 遍历索引数组，并为每个索引添加到tableInfo中的每个对象
      const result = indexes.map((index, indexIndex) => {
        // 遍历tableInfo数组
        const info = that.tableInfo[indexIndex];
        // 如果没有找到对应的索引，返回一个空对象
        if (!info) {
          return {'index':index,...that.tableObj};
        }
        // 添加index属性
        info.index = index;
        // 返回修改后的对象
        return info;
      });
      console.log(result);
      return result
    },
    generate() {
      this.tableInfo = this.addIndex18()
      console.log('tableInfo:',this.tableInfo);
      let that = this;
      // 读取并获得模板文件的二进制内容
      JSZipUtils.getBinaryContent(
        "./alertTemplateA.docx",
        function (error, content) {
          // model.docx是模板。我们在导出的时候，会根据此模板来导出对应的数据
          // 抛出异常
          if (error) {
            throw error;
          }

          // 创建一个PizZip实例，内容为模板的内容
          let zip = new PizZip(content);
          // 创建并加载docxtemplater实例对象
          let doc = new docxtemplater().loadZip(zip);
          // 设置模板变量的值
          doc.setData({
            header: that.fields.header,
            basicInfo: that.fieldsbasicInfo,
            tableInfo: that.fieldstableInfo,
          });
          try {
            // 用模板变量的值替换所有模板变量
            doc.render();
          } catch (error) {
            // 抛出异常
            let e = {
              message: error.message,
              name: error.name,
              stack: error.stack,
              properties: error.properties,
            };
            console.log(JSON.stringify({ error: e }));
            throw error;
          }
          // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
          let out = doc.getZip().generate({
            type: "blob",
            mimeType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          // 将目标文件对象保存为目标类型的文件，并命名
          saveAs(out, "testA.docx");
        }
      );
    },
  },
};
</script>
