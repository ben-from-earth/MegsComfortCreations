export function isThreeDsConsole(consoleName: string | null | undefined) {
  return consoleName?.trim().toLowerCase() === '3ds';
}
