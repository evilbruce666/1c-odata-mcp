/** Сегодня по часам компьютера, а не по UTC: ночью в России дата в UTC ещё вчерашняя. */
export function localDate(d: Date): string {
  const two = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())}`;
}
