const AI_PHRASES: RegExp[] = [
  /^希望这[对你].+[！!]\s*$/gim,
  /^如果.+(?:问题|需要|帮助).+请.+(?:告诉|联系|随时).+[！!。.]\s*$/gim,
  /^Would you like me to.+[?]?\s*$/gim,
  /^Let me know if.+[!.]?\s*$/gim,
  /^I hope this helps[!.]?\s*$/gim,
  /^Hope (?:this|that) helps[!.]?\s*$/gim,
  /^As an AI(?: language model)?,.+[!.]\s*$/gim,
  /^当然[，,]很[高开]兴.+[！!]\s*$/gim,
  /^Sure[,!].+/gim,
  /^Certainly[!.,]/gim,
  /^如果有(?:任何)?(?:更多|其他)(?:问题|疑问).+/gim,
  /^以上[，,]希望.+帮助[。.]?\s*$/gim,
  /^请注意[：:].+/gim,
  /^最后[，,].+提醒[：:].+/gim,
  /^重要提示[：:].+/gim,
  /^免责声明[：:][\s\S]*?(?=\n\n|$)/gim,
  /^【温馨提示】[\s\S]*?(?=\n\n|$)/gim,
]

export function cleanAIPhrases(text: string): string {
  let result = text
  for (const pattern of AI_PHRASES) {
    result = result.replace(pattern, '')
  }
  return result.replace(/\n{3,}/g, '\n\n').trim()
}
