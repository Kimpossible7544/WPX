export default {
  async fetch(request, env) {
    const corsHeaders = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, X-WPX-Secret',
    };

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: corsHeaders });
    }

    const secret = request.headers.get('X-WPX-Secret');
    if (secret !== 'Abigail2011!') {
      return new Response(JSON.stringify({ error: 'Unauthorized' }), {
        status: 401,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });
    }

    const url = new URL(request.url);
    const action = url.searchParams.get('action');
    const key = url.searchParams.get('key');

    try {
      if (action === 'ai') {
        const body = await request.json();
        const resp = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': env.ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01'
          },
          body: JSON.stringify(body)
        });
        const data = await resp.json();
        return new Response(JSON.stringify(data), {
          status: resp.status,
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      }

      const KV = env.WPX_Data || env.WPX_DATA || env.KV;

      if (request.method === 'GET') {
        if (!key) return new Response(JSON.stringify({ error: 'No key' }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
        const value = await KV.get(key);
        return new Response(JSON.stringify({ value }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      }

      if (request.method === 'POST') {
        if (!key) return new Response(JSON.stringify({ error: 'No key' }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
        const body = await request.json();
        await KV.put(key, JSON.stringify(body.value));
        return new Response(JSON.stringify({ success: true }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      }

      if (request.method === 'DELETE') {
        if (!key) return new Response(JSON.stringify({ error: 'No key' }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
        await KV.delete(key);
        return new Response(JSON.stringify({ success: true }), {
          headers: { ...corsHeaders, 'Content-Type': 'application/json' }
        });
      }

      return new Response(JSON.stringify({ error: 'Method not allowed' }), {
        status: 405,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });

    } catch (err) {
      return new Response(JSON.stringify({ error: err.message }), {
        status: 500,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' }
      });
    }
  }
};
