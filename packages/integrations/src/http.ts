export class HttpError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly responseBody: string
  ) {
    super(message);
  }
}

export async function fetchJson<T>(
  input: string,
  init: RequestInit = {},
  fetcher: typeof fetch = fetch
): Promise<T> {
  const response = await fetcher(input, init);
  if (!response.ok) {
    const responseBody = await response.text();
    throw new HttpError(`${init.method ?? "GET"} ${input} failed`, response.status, responseBody);
  }
  return (await response.json()) as T;
}
