// tslint:disable-next-line:export-name
export default class Utility {
    public static formatDate(d: Date): string {
        return `${this.pad(d.getDate(), 2)}.${this.pad(d.getMonth() + 1, 2)}.${d.getFullYear().toString()}`;
    }

    public static pad(n: number, zeroes: number): string {
        let s: string = n.toString();
        while (s.length < zeroes) {
            s = `0${s}`;
        }
        return s;
    }

    public static endsWith(s: string, e: string): boolean {
        if (s && s.length && e && e.length) {
            return s.lastIndexOf(e) === s.length - e.length;
        }
        return false;
    }

    public static trimEnd(s: string, e: string): string {
        if (s && s.length && e && e.length) {
            while (Utility.endsWith(s, e)) {
                s = s.substring(0, s.length - e.length);
            }
        }
        return s;
    }
}