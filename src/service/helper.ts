export function timeAgo(input: Date | string | number): string {
  const d = new Date(input);
  const now = new Date();
  const diffMs = now.getTime() - d.getTime();

  const diffSec = Math.floor(diffMs / 1000);
  const diffMin = Math.floor(diffSec / 60);
  const diffHrs = Math.floor(diffMin / 60);
  const diffDays = Math.floor(diffHrs / 24);

  if (diffDays === 0) {
    if (diffHrs > 0) return `${diffHrs} hour${diffHrs > 1 ? "s" : ""} ago`;
    if (diffMin > 0) return `${diffMin} minute${diffMin > 1 ? "s" : ""} ago`;
    return "Just now";
  }

  if (diffDays === 1) return "Yesterday";
  if (diffDays < 7) return `${diffDays} day${diffDays > 1 ? "s" : ""} ago`;

  return d.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}
