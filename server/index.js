import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import express from 'express';
import { readFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

function loadSkill() {
  try {
    const skillPath = join(__dirname, '..', 'skills', 'ariba-crd-chargeability', 'SKILL.md');
    const content = readFileSync(skillPath, 'utf-8');
    const frontmatter = content.match(/^---\n([\s\S]*?)\n---/);
    let version = '1.0.0';
    let name = 'ariba-crd-chargeability';
    if (frontmatter) {
      const versionMatch = frontmatter[1].match(/version:\s*(.+)/);
      const nameMatch = frontmatter[1].match(/name:\s*(.+)/);
      if (versionMatch) version = versionMatch[1].trim();
      if (nameMatch) name = nameMatch[1].trim();
    }
    return { name, version, content };
  } catch (e) {
    return { error: 'Skill not found', message: e.message };
  }
}

const server = new McpServer({ name: 'ariba-crd-skill-server', version: '1.0.0' });

server.tool(
  'get_crd_chargeability_skill',
  'Returns the latest Ariba CRD chargeability skill instructions and rules',
  {},
  async () => {
    const skill = loadSkill();
    return { content: [{ type: 'text', text: skill.content }] };
  }
);

app.post('/mcp', async (req, res) => {
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  res.on('close', () => transport.close());
  await server.connect(transport);
  await transport.handleRequest(req, res, req.body);
});

app.get('/health', (req, res) => {
  const skill = loadSkill();
  res.json({ status: 'ok', skill: skill.name, version: skill.version });
});

app.listen(PORT, () => {
  const skill = loadSkill();
  console.log(`MCP Server running on port ${PORT} | Skill: ${skill.name} v${skill.version}`);
});
