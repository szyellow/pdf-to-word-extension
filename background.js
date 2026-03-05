// 后台脚本 - 处理图标点击事件
chrome.action.onClicked.addListener(async (tab) => {
  try {
    await chrome.tabs.sendMessage(tab.id, { action: 'togglePanel' });
  } catch (error) {
    console.error('发送消息失败:', error);
  }
});

chrome.runtime.onInstalled.addListener(() => {
  console.log('PDF转Word插件已安装');
});
