
import { getCollection } from "astro:content";
import rss from "@astrojs/rss";
import { SITE_DESCRIPTION, SITE_TITLE } from "../consts";

export async function GET(context) {
  const posts = await getCollection("blog");
  const essays = await getCollection("essays");

  const items = [
    ...posts.map((post) => ({
      ...post.data,
      link: `${import.meta.env.BASE_URL}/blog/${post.id}/`,
    })),
    ...essays.map((essay) => ({
      ...essay.data,
      link: `${import.meta.env.BASE_URL}/essays/${essay.id}/`,
    })),
  ].sort((a, b) => b.pubDate.getTime() - a.pubDate.getTime());

  return rss({
    title: SITE_TITLE,
    description: SITE_DESCRIPTION,
    site: new URL(import.meta.env.BASE_URL, context.site).toString(),
    items,
  });
}
