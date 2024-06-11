<template>
  <div>
    <!-- {{ fields }} -->
    <h1>报警器检验检测原始记录封面</h1>
    <Form
      :model="fields"
      :label-width="150"
      :rules="ruleCustom"
      style="width: 60%; margin: 5% 20%">
      <FormItem label="记录年份：" prop="recordYear">
        <Input
          v-model="fields.header[0].recordYear"
          placeholder="输入记录年份..."></Input>
      </FormItem>
      <FormItem label="记录编号：" prop="recordNo">
        <Input
          v-model="fields.header[0].recordNo"
          placeholder="输入记录编号..."></Input>
      </FormItem>
      <FormItem label="当前页码：" prop="currPage">
        <Input
          v-model="fields.header[0].currPage"
          placeholder="输入当前页码..."></Input>
      </FormItem>
      <FormItem label="总页数：" prop="totalPage">
        <Input
          v-model="fields.header[0].totalPage"
          placeholder="输入总页数..."></Input>
      </FormItem>
      <FormItem label="公司：" prop="company">
        <Input
          v-model="fields.basicInfo[0].company"
          placeholder="输入公司名称..."></Input>
      </FormItem>
      <FormItem label="温度：" prop="temperature">
        <Input
          v-model="fields.basicInfo[0].temperature"
          placeholder="输入温度..."></Input>
      </FormItem>
      <FormItem label="湿度：" prop="humidity">
        <Input
          v-model="fields.basicInfo[0].humidity"
          placeholder="输入湿度..."></Input>
      </FormItem>
      <!-- 可以添加的动态增减表单项 -->
      <FormItem label="气体检测报警仪基本信息：" prop="dynamicItems">
        <div v-for="(item, index) in fields.tableInfo" :key="index">
          <Row :key="index">
            <Col span="6">
              <!-- Sample Name 输入框 -->
              <FormItem
                :label="'样品名称 ' + (index + 1)"
                :prop="'tableInfo.' + index + '.sampleName'">
                <Input
                  v-model="item.sampleName"
                  :placeholder="'请输入样品名称' + (index + 1)"></Input>
              </FormItem>
            </Col>
            <Col span="6">
              <!-- Factory 输入框 -->
              <FormItem
                :label="'制造厂名称 ' + (index + 1)"
                :prop="'tableInfo.' + index + '.factory'">
                <Input
                  v-model="item.factory"
                  :placeholder="'请输入制造厂名称 ' + (index + 1)"></Input>
              </FormItem>
            </Col>
            <Col span="6">
              <!-- Target Gas 输入框 -->
              <FormItem
                :label="'探测气体 ' + (index + 1)"
                :prop="'tableInfo.' + index + '.targetGas'">
                <Input
                  v-model="item.targetGas"
                  :placeholder="'请输入探测气体名称 ' + (index + 1)"></Input>
              </FormItem>
            </Col>
            <Col span="5">
              <!-- Scale Range 输入框 -->
              <FormItem
                :label="'量程范围 ' + (index + 1)"
                :prop="'tableInfo.' + index + '.scaleRange'">
                <Input
                  v-model="item.scaleRange"
                  :placeholder="'请输入量程范围 ' + (index + 1)"></Input>
              </FormItem>
            </Col>
            <Col span="1">
              <!-- 删除按钮 -->
              <Button type="error" @click="removeDynamicFormItem(index)"
                >Delete</Button
              >
            </Col>
          </Row>
        </div>
        <Button type="dashed" @click="addDynamicFormItem" long style="margin-top: 10px;">添加一项</Button>
      </FormItem>
    </Form>
    <Button type="primary" @click="generate" long icon="ios-download-outline" size="large">生成模版A</Button>
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
      fields: {
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
            sampleName: "LD-200S固定式气体探测器",
            factory: "南京更佳电子科技有限公司",
            targetGas: "可燃气",
            scaleRange: "0-100％LEL",
          },
          {
            sampleName: "LD-200S固定式气体探测器",
            factory: "南京更佳电子科技有限公司",
            targetGas: "H2S",
            scaleRange: "0-100ppm",
          },
        ],
      },
      tableObj: {
        sampleName: "",
        factory: "",
        targetGas: "",
        scaleRange: "",
      },
    };
  },
  computed: {},
  watch: {},

  mounted() {},
  methods: {
    removeDynamicFormItem(index) {
      this.fields.tableInfo.splice(index, 1);
    },
    addDynamicFormItem() {
      if (this.fields.tableInfo.length < 18) {
      this.fields.tableInfo.push({
        sampleName: "",
        factory: "",
        targetGas: "",
        scaleRange: "",
      });
    }else{
      this.$Message.error('最多支持18项，请另起一页填写');
    }
    },

    addIndex18() {
      let that = this;
      const indexes = Array.from("abcdefghijklmnopqr");
      const result = indexes.map((index, indexIndex) => {
        const info = that.fields.tableInfo[indexIndex];
        if (!info) {
          return { ...that.tableObj, index };
        }
        info.index = index;
        return info;
      });
      console.log(result);
      return result;
    },
    generate() {
      this.fields.tableInfo = this.addIndex18();
      console.log("tableInfo:", this.fields.tableInfo);
      let that = this;
      JSZipUtils.getBinaryContent(
        "./alertTemplateA.docx",
        function (error, content) {
          if (error) {
            throw error;
          }
          let zip = new PizZip(content);
          let doc = new docxtemplater().loadZip(zip);
          doc.setData({
            header: that.fields.header,
            basicInfo: that.fields.basicInfo,
            tableInfo: that.fields.tableInfo,
          });
          try {
            doc.render();
          } catch (error) {
            let e = {
              message: error.message,
              name: error.name,
              stack: error.stack,
              properties: error.properties,
            };
            console.log(JSON.stringify({ error: e }));
            throw error;
          }
          let out = doc.getZip().generate({
            type: "blob",
            mimeType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          saveAs(out, "testA.docx");
        }
      );
    },

    // generate() {
    //   this.tableInfo = this.addIndex18();
    //   console.log("tableInfo:", this.tableInfo);
    //   let that = this;
    //   // 读取并获得模板文件的二进制内容
    //   JSZipUtils.getBinaryContent(
    //     "./alertTemplateA.docx",
    //     function (error, content) {
    //       // model.docx是模板。我们在导出的时候，会根据此模板来导出对应的数据
    //       // 抛出异常
    //       if (error) {
    //         throw error;
    //       }

    //       // 创建一个PizZip实例，内容为模板的内容
    //       let zip = new PizZip(content);
    //       // 创建并加载docxtemplater实例对象
    //       let doc = new docxtemplater().loadZip(zip);
    //       // 设置模板变量的值
    //       doc.setData({
    //         header: that.fields.header,
    //         basicInfo: that.fieldsbasicInfo,
    //         tableInfo: that.fieldstableInfo,
    //       });
    //       try {
    //         // 用模板变量的值替换所有模板变量
    //         doc.render();
    //       } catch (error) {
    //         // 抛出异常
    //         let e = {
    //           message: error.message,
    //           name: error.name,
    //           stack: error.stack,
    //           properties: error.properties,
    //         };
    //         console.log(JSON.stringify({ error: e }));
    //         throw error;
    //       }
    //       // 生成一个代表docxtemplater对象的zip文件（不是一个真实的文件，而是在内存中的表示）
    //       let out = doc.getZip().generate({
    //         type: "blob",
    //         mimeType:
    //           "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    //       });
    //       // 将目标文件对象保存为目标类型的文件，并命名
    //       saveAs(out, "testA.docx");
    //     }
    //   );
    // },
  },
};
</script>
