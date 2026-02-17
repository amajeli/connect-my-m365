import { getUncachableStripeClient } from './stripeClient';

async function createProducts() {
  const stripe = await getUncachableStripeClient();

  const products = await stripe.products.search({ query: "name:'Outlook Connector Pro'" });
  if (products.data.length > 0) {
    console.log('Product already exists:', products.data[0].id);
    const prices = await stripe.prices.list({ product: products.data[0].id, active: true });
    console.log('Prices:', prices.data.map(p => `${p.id}: $${(p.unit_amount || 0) / 100}/${p.recurring?.interval}`));
    return;
  }

  const product = await stripe.products.create({
    name: 'Outlook Connector Pro',
    description: 'Connect your Outlook email and calendar to Claude.ai via MCP. Get real-time access to your emails and calendar events directly in Claude conversations.',
    metadata: {
      type: 'subscription',
    },
  });

  const monthlyPrice = await stripe.prices.create({
    product: product.id,
    unit_amount: 999,
    currency: 'usd',
    recurring: { interval: 'month' },
  });

  const yearlyPrice = await stripe.prices.create({
    product: product.id,
    unit_amount: 7999,
    currency: 'usd',
    recurring: { interval: 'year' },
  });

  console.log('Created product:', product.id);
  console.log('Monthly price:', monthlyPrice.id, '- $9.99/month');
  console.log('Yearly price:', yearlyPrice.id, '- $79.99/year');
}

createProducts().catch(console.error);
