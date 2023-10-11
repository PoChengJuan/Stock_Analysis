Public Enum enuRabbitMQ '等同Routing Key
  'HOST 交互
  HOST_TO_WMS
  WMS_TO_HOST

  'GUI 交互
  GUI_TO_HOST
  HOST_TO_GUI

  'MCS 交互
  MCS_TO_HOST
  HOST_TO_MCS

  'NS 交互
  NS_TO_HOST
  HOST_TO_NS
End Enum


Public Enum enuMQHeaders
  '一定要有的
  FUNCTION_ID
  UUID
  SEND_SYSTEM 'WMS:1/MCS:3/GUI:4
  RECEIVE_SYSTEM 'WMS:1/MCS:3/GUI:4
  DIRECTION '發送的給 Primary/ 回覆的給 Secondary
  USER_ID
  CLIENT_ID
  IP
  CREATE_TIME

  '回覆時使用
  RESULT
  RESULT_MESSAGE

End Enum