<template>
  <div>
    <p>Name: {{ profile?.displayName }}</p>
    <p>Title: {{ profile?.jobTitle }}</p>
    <p>Mail: {{ profile?.mail }}</p>
    <p>
      Phone:
      {{ profile?.businessPhones ? profile?.businessPhones[0] : '' }}
    </p>
    <p>Location: {{ profile?.officeLocation }}</p>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue';
import { useMsGraph } from '@/composition-api/useMsGraph';

const { getGraphProfile } = useMsGraph();

const profile = ref();

onMounted(async () => {
  profile.value = await getGraphProfile();
});
</script>
