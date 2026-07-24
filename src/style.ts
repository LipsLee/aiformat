export interface StyleConfig {
  fontFamily: string
  fontSize: number
  lineHeight: number
  paraSpacing: number
  textColor: string
}

export function getStyleConfig(): StyleConfig {
  return {
    fontFamily: (document.getElementById('fontFamily') as HTMLSelectElement).value,
    fontSize: parseInt((document.getElementById('fontSize') as HTMLInputElement).value),
    lineHeight: parseFloat((document.getElementById('lineHeight') as HTMLInputElement).value),
    paraSpacing: parseInt((document.getElementById('paraSpacing') as HTMLInputElement).value),
    textColor: (document.getElementById('textColor') as HTMLInputElement)?.value || '#1a1a1a',
  }
}
