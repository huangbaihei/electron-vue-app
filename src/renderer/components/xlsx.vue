<template>
  <div class="xlsx">
    <div class="form-item">
      <div class="form-title">1、确认excel表列字段名和顺序</div>
      <div class="form-input">
        <div>账户号 交易描述 卡号 交易日期 币种 交易金额 余额 名字 证件号码 编号</div>
      </div>
    </div>
    <!-- <div class="form-item">
      <div class="form-title">2、手动拆分excel表成多个文件</div>
      <div class="form-input">
        <div>每个文件不能大于30M，30M大约40万行</div>
      </div>
    </div> -->
    <div class="form-item">
      <div class="form-title">2、选择需要处理的excel文件</div>
      <div class="form-input">
        <el-upload
          ref="elUploadFile"
          drag
          multiple
          action=""
          :auto-upload="false"
          :on-change="onFileChange"
          :on-remove="onFileRemove">
          <i class="el-icon-upload"></i>
          <div class="el-upload__text">
            <div>将文件拖到此处</div>
            <div>(支持多个)</div>
          </div>
        </el-upload>
      </div>
    </div>
    <div class="form-item">
      <div class="form-title">3、选择放置输出excel文件的文件夹</div>
      <div class="form-input">
        <el-upload
          class="el-upload-folder"
          ref="elUploadFolder"
          drag
          action=""
          :limit="1"
          :auto-upload="false"
          :on-change="onFolderChange"
          :on-remove="onFolderRemove"
          list-type="picture">
          <i class="el-icon-upload"></i>
          <div class="el-upload__text">将文件夹拖到此处</div>
        </el-upload>
      </div>
    </div>
    <el-button type="primary" round @click="toStart" :loading="isLoading">{{isLoading ? '正在处理' : '开始处理'}}</el-button>
    <div class="log-text" v-if="!log.error">{{log.text}}</div>
    <div class="log-error" v-else>{{log.error}}</div>
  </div>
</template>

<script>
import { Button, Upload, Input, Message } from 'element-ui'
import splitXlsx from './app-xlsx'
const { shell } = require('electron')
const os = require('os')
const net = require('net')

export default {
  components: {
    [Button.name]: Button,
    [Upload.name]: Upload,
    [Input.name]: Input
  },
  data () {
    return {
      sourcePathList: [],
      folderPath: '',
      isLoading: false,
      log: {
        text: '',
        error: ''
      }
    }
  },
  mounted () {
    this.$refs.elUploadFile.$refs['upload-inner'].$refs.input.onclick = (e) => {
      e.preventDefault()
      this.showItemInFolder()
    }
    this.$refs.elUploadFolder.$refs['upload-inner'].$refs.input.onclick = (e) => {
      e.preventDefault()
      this.showItemInFolder()
    }
    const html = document.querySelector('html')
    html.ondragenter = html.ondragover = html.ondragleave = html.ondrop = e => {
      e.preventDefault()
    }
  },
  methods: {
    toStart () {
      this.log = {
        text: '',
        error: ''
      }
      let msg = ''
      if (!this.sourcePathList.length) {
        msg = '请选择需要处理的excel文件'
      } else if (!this.folderPath) {
        msg = '请选择放置输出excel文件的文件夹'
      }
      if (msg) {
        Message.warning({
          message: msg,
          center: true
        })
        return false
      }
      this.isLoading = true
      // this.log.text = '文件太大的话可能需要等待3-5分钟...'
      setTimeout(async () => {
        // for (let i = 0; i < this.sourcePathList.length; i++) {
        //   await splitXlsx(
        //     this.sourcePathList[i],
        //     i + 1,
        //     this.sourcePathList.length,
        //     this.folderPath,
        //     this.log
        //   ).catch(e => {
        //     this.log.error = e
        //   })
        // }
        // this.isLoading = false

        let tcpData = ''
        const client = net.connect({ port: 9800 })
        // client.write(JSON.stringify(this.sourcePathList))
        client.write(JSON.stringify([...this.sourcePathList, this.folderPath]))
        client.on('data', async data => {
          tcpData += data.toString()
          if (tcpData.startsWith('[')) {
            // 第一个通信会回抛数组
            if (tcpData.endsWith('数据传输结束标记')) {
              data = JSON.parse(tcpData.replace(/数据传输结束标记$/, ''))
              for (let i = 0; i < data.length; i++) {
                await splitXlsx(
                  data[i],
                  i + 1,
                  data.length,
                  this.folderPath,
                  this.log,
                  this.sourcePathList.length
                ).catch(e => {
                  this.log.error = e
                  this.isLoading = false
                })
              }
              tcpData = ''
              // 处理完以后，传文件夹过去给go批量生成xlsx格式
              client.write(this.folderPath)
            }
          } else {
            // 之后的通信都是回抛处理状态
            if (tcpData.startsWith('处理中断')) {
              this.log.error = tcpData
              this.isLoading = false
            } else {
              this.log.text = tcpData
              if (this.log.text.startsWith('处理结束')) {
                this.isLoading = false
              }
            }
            tcpData = ''
          }
        })
      }, 1)
    },
    showItemInFolder () {
      shell.showItemInFolder(os.homedir())
    },
    onFileChange (file, fileList) {
      this.sourcePathList = fileList.map(x => x.raw.path)
    },
    onFileRemove (file, fileList) {
      this.sourcePathList = fileList.map(x => x.raw.path)
    },
    onFolderChange (file, fileList) {
      fileList[0].url = require('@/assets/icon-folder.png')
      this.folderPath = file.raw.path
    },
    onFolderRemove (file, fileList) {
      this.folderPath = ''
    }
  }
}
</script>

<style lang="scss" scoped>
  .xlsx {
    height: 100vh;
    padding: 0 50px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    overflow: auto;
  }
  .flex {
    display: flex;
    align-items: center;
  }
  .form-item {
    display: flex;
    margin-bottom: 30px;
    .form-title {
      width: 265px;
      flex-shrink: 0;
    }
    .form-input {
      width: 500px;
      flex-shrink: 0;
    }
    .folder-path {
      width: 360px;
      flex-shrink: 0;
    }
  }
  .el-upload-folder {
    /deep/.el-upload-list__item {
      height: auto;
      padding: 0 10px 0 90px;
      display: flex;
      align-items: center;
      .el-upload-list__item-thumbnail {
        height: 15px;
        width: 15px;
      }
      .el-icon-document {
        display: none;
      }
      .el-upload-list__item-name {
        margin: 0;
        margin-left: 8px;
        padding: 0;
      }
      .el-icon-close {
        top: 50%;
        transform: translateY(-50%);
      }
    }
  }
  /deep/.el-upload-list {
    width: 360px;
  }
  .log-text {
    margin-top: 10px;
    color: green;
  }
  .log-error {
    margin-top: 10px;
    color: red;
  }
</style>