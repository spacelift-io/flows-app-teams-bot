/**
 * Block Registry for Teams Bot
 *
 * This file exports all blocks as a dictionary for easy registration.
 */

// Messaging blocks
import { sendMessage } from "./messaging/sendMessage.ts";
import { updateMessage } from "./messaging/updateMessage.ts";
import { deleteMessage } from "./messaging/deleteMessage.ts";

// Subscription blocks
import { mentionsSubscription } from "./subscriptions/mentionsSubscription.ts";
import { subscribeToReplies } from "./subscriptions/subscribeToReplies.ts";

// User blocks
import { getUserInfo } from "./users/getUserInfo.ts";

/**
 * Dictionary of all available blocks
 */
export const blocks = {
  // Messaging
  sendMessage,
  updateMessage,
  deleteMessage,

  // Subscriptions
  mentionsSubscription,
  threadSubscription: subscribeToReplies,

  // Users
  getUserInfo,
} as const;

// Named exports for individual blocks
export {
  sendMessage,
  updateMessage,
  deleteMessage,
  mentionsSubscription,
  subscribeToReplies as threadSubscription,
  getUserInfo,
};
