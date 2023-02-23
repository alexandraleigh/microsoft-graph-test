<template>
  <div class="home">
    <span v-if="!isAuthenticated">Please sign-in to see your profile information.</span>

    <div v-else>
      <el-button type="primary" @click="goTo('profile')">Request Profile Information</el-button>
      <el-button type="primary" @click="goTo('drive-files')">Request Drive Files</el-button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { useRouter } from "vue-router";
import { useIsAuthenticated } from "../composition-api/useIsAuthenticated";

const isAuthenticated = useIsAuthenticated();

const router = useRouter();

function goTo(page) {
  router.push(`/${page}`);
}
</script>
