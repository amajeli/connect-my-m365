// Netlify Function that runs the whole Express API (auth, OAuth, MCP, billing).
// netlify.toml routes /api/*, /oauth/*, /mcp/* and /.well-known/* here; the
// React client is served from the CDN.
import serverless from "serverless-http";
import { createApp } from "../../server/app";

let handlerPromise: Promise<ReturnType<typeof serverless>> | undefined;

export const handler = async (event: any, context: any) => {
  handlerPromise ??= createApp().then((app) => serverless(app));
  return (await handlerPromise)(event, context);
};
